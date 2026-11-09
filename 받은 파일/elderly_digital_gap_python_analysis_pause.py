# -*- coding: utf-8 -*-
"""
고령자 디지털 소외 위험 예측 - Weka vs Python 비교분석
======================================================

목적
----
1) NIA 2025 디지털정보격차 실태조사 고령층 ARFF 데이터를 Python에서 재현 분석
2) Weka와 동일하게 10-fold 교차검증을 사용
3) ZeroR / Logistic / Decision Tree(CART) / RandomForest 성능 평가
4) Weka 결과와 Python 결과 비교
5) Logistic 오즈비(OR)와 RandomForest 변수 중요도 확인
6) 발표/보고서용 CSV, PNG, TXT 결과 자동 저장

중요한 해석 주의
----------------
- Weka와 Python 모두 10-fold 교차검증을 사용하지만, 두 프로그램이 내부적으로
  완전히 동일한 fold 구성을 만들었다고 보장할 수는 없습니다.
- Weka J48은 C4.5 계열이고, sklearn DecisionTreeClassifier는 CART 계열입니다.
  따라서 두 의사결정나무의 성능 차이는 단순히 "도구의 우열"로 해석하면 안 됩니다.
- Python 결과는 Weka 결과를 억지로 맞추기 위한 튜닝이 아니라,
  Python의 대표적인 기본 구현으로 재현한 비교 결과입니다.

실행 예
-------
python elderly_digital_gap_python_analysis.py
python elderly_digital_gap_python_analysis.py --input "2025 정보격차 고령층.arff"
python elderly_digital_gap_python_analysis.py --input "2025 정보격차 고령층.arff" --output "python_results"

필요 패키지
-----------
pip install numpy pandas scipy scikit-learn matplotlib
"""

import argparse
import sys
from pathlib import Path
import warnings

import numpy as np
import pandas as pd
from scipy.io import arff

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.pipeline import Pipeline
from sklearn.model_selection import StratifiedKFold, cross_val_predict
from sklearn.dummy import DummyClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    cohen_kappa_score,
    roc_auc_score,
    confusion_matrix,
)

warnings.filterwarnings("ignore", category=FutureWarning)

# ---------------------------------------------------------------------
# 0. 기본 설정
# ---------------------------------------------------------------------

DEFAULT_FILENAME = "2025 정보격차 고령층.arff"
TARGET = "risk_class"
NUMERIC_COLS = ["confidence"]

EXPECTED_FEATURES = [
    "age_group",
    "region",
    "area_type",
    "income",
    "education",
    "sex",
    "pc_available",
    "smartphone",
    "confidence",
    "training_experience",
]

# 팀이 Weka에서 정리한 결과
# 출처: 현재 프로젝트 정리본
WEKA_RESULTS = {
    "ZeroR": {
        "Accuracy": 0.688,
        "Kappa": 0.000,
        "ROC_AUC": 0.498,
        "Risk_Recall": 1.000,
    },
    "Logistic": {
        "Accuracy": 0.819,
        "Kappa": 0.568,
        "ROC_AUC": 0.896,
        "Risk_Recall": 0.889,
    },
    "J48": {
        "Accuracy": 0.814,
        "Kappa": 0.567,
        "ROC_AUC": 0.855,
        "Risk_Recall": 0.866,
    },
    "RandomForest": {
        "Accuracy": 0.799,
        "Kappa": 0.524,
        "ROC_AUC": 0.871,
        "Risk_Recall": 0.867,
    },
}

# Logistic OR 해석 시 사용할 기준범주(reference category)
# 첫 번째 값이 기준범주가 됩니다.
CATEGORY_ORDERS = {
    "age_group": ["50s", "60s", "70plus"],
    "region": [
        "Seoul", "Busan", "Daegu", "Incheon", "Gwangju", "Daejeon",
        "Ulsan", "Sejong", "Gyeonggi", "Gangwon", "Chungbuk", "Chungnam",
        "Jeonbuk", "Jeonnam", "Gyeongbuk", "Gyeongnam", "Jeju"
    ],
    "area_type": ["dong", "eupmyeon"],
    "income": ["under200", "200to399", "400to599", "600plus"],
    "education": ["college", "high", "middle", "elementary"],
    "sex": ["male", "female"],
    "pc_available": ["yes", "no"],
    "smartphone": ["yes", "no"],
    "training_experience": ["yes", "no"],
}


# ---------------------------------------------------------------------
# 1. 유틸리티
# ---------------------------------------------------------------------

def make_ohe(drop=None, categories="auto"):
    """scikit-learn 버전 차이를 고려해 OneHotEncoder 생성."""
    try:
        return OneHotEncoder(
            handle_unknown="ignore",
            drop=drop,
            categories=categories,
            sparse_output=False,
        )
    except TypeError:
        return OneHotEncoder(
            handle_unknown="ignore",
            drop=drop,
            categories=categories,
            sparse=False,
        )


def find_default_input():
    """입력 파일 경로가 지정되지 않았을 때 흔한 위치에서 자동 탐색."""
    script_dir = Path(__file__).resolve().parent
    candidates = [
        Path.cwd() / DEFAULT_FILENAME,
        script_dir / DEFAULT_FILENAME,
        Path("/mnt/data") / DEFAULT_FILENAME,
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def decode_bytes_columns(df):
    """scipy ARFF 로더가 byte 문자열로 읽은 값을 일반 문자열로 변환."""
    out = df.copy()
    for col in out.select_dtypes(include="object").columns:
        out[col] = out[col].apply(
            lambda x: x.decode("utf-8") if isinstance(x, (bytes, bytearray)) else x
        )
    return out


def save_csv(df, path):
    """한국어 Windows Excel에서도 잘 열리도록 UTF-8-SIG 저장."""
    df.to_csv(path, index=False, encoding="utf-8-sig")


def pct(x):
    return f"{x * 100:.2f}%"


# ---------------------------------------------------------------------
# 2. 데이터 로드 / 검증
# ---------------------------------------------------------------------

def load_and_validate(input_path):
    raw, meta = arff.loadarff(str(input_path))
    df = pd.DataFrame(raw)
    df = decode_bytes_columns(df)

    missing_cols = [c for c in EXPECTED_FEATURES + [TARGET] if c not in df.columns]
    if missing_cols:
        raise ValueError(
            "필수 컬럼이 없습니다: " + ", ".join(missing_cols)
        )

    # 이번 프로젝트에서 사용하는 10개 입력변수 + 종속변수만 유지
    df = df[EXPECTED_FEATURES + [TARGET]].copy()

    # 목표변수 확인
    valid_classes = {"risk", "nonrisk"}
    actual_classes = set(df[TARGET].dropna().unique())
    if not actual_classes.issubset(valid_classes):
        raise ValueError(
            f"{TARGET} 값이 예상과 다릅니다. 현재 값: {sorted(actual_classes)}"
        )

    # 결측 확인
    missing_count = int(df.isna().sum().sum())
    if missing_count > 0:
        print(f"[주의] 결측값이 총 {missing_count}개 있습니다.")
        print(df.isna().sum()[df.isna().sum() > 0])
        print("이번 스크립트는 결측이 없는 현재 ARFF를 기준으로 설계되었습니다.")

    return df


# ---------------------------------------------------------------------
# 3. 기술통계 / 그룹별 소외율
# ---------------------------------------------------------------------

def make_descriptive_outputs(df, output_dir):
    work = df.copy()
    work["risk_binary"] = (work[TARGET] == "risk").astype(int)

    class_dist = (
        work[TARGET]
        .value_counts()
        .rename_axis("class")
        .reset_index(name="count")
    )
    class_dist["ratio"] = class_dist["count"] / len(work)
    save_csv(class_dist, output_dir / "01_class_distribution.csv")

    group_rows = []
    group_vars = [
        "age_group",
        "education",
        "income",
        "area_type",
        "sex",
        "pc_available",
        "smartphone",
        "training_experience",
        "region",
    ]

    for col in group_vars:
        g = (
            work.groupby(col, dropna=False)["risk_binary"]
            .agg(["size", "sum", "mean"])
            .reset_index()
            .rename(
                columns={
                    col: "category",
                    "size": "n",
                    "sum": "risk_count",
                    "mean": "risk_rate",
                }
            )
        )
        g.insert(0, "variable", col)
        group_rows.append(g)

    group_rates = pd.concat(group_rows, ignore_index=True)
    save_csv(group_rates, output_dir / "02_group_risk_rates.csv")

    # 발표용 대표 그룹
    for col in ["age_group", "education"]:
        sub = group_rates[group_rates["variable"] == col].copy()
        sub["risk_rate_pct"] = sub["risk_rate"] * 100

        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(sub["category"], sub["risk_rate_pct"])
        ax.set_title(f"Risk rate by {col}")
        ax.set_xlabel(col)
        ax.set_ylabel("Risk rate (%)")
        ax.set_ylim(0, 100)
        ax.grid(axis="y", alpha=0.25)
        fig.tight_layout()
        fig.savefig(output_dir / f"chart_risk_rate_{col}.png", dpi=180)
        plt.close(fig)

    return class_dist, group_rates


# ---------------------------------------------------------------------
# 4. 10-fold 교차검증 모델 비교
# ---------------------------------------------------------------------

def build_common_preprocessor(X):
    categorical_cols = [c for c in X.columns if c not in NUMERIC_COLS]

    preprocessor = ColumnTransformer(
        transformers=[
            (
                "cat",
                make_ohe(drop=None),
                categorical_cols,
            ),
            (
                "num",
                "passthrough",
                NUMERIC_COLS,
            ),
        ],
        remainder="drop",
    )
    return preprocessor


def evaluate_models(df, output_dir):
    X = df[EXPECTED_FEATURES].copy()
    y = (df[TARGET] == "risk").astype(int)

    # Weka의 10-fold Cross-validation과 평가방식을 맞추기 위한 Python 측 설정.
    # risk/nonrisk 비율을 각 fold에서 최대한 유지하기 위해 StratifiedKFold 사용.
    cv = StratifiedKFold(
        n_splits=10,
        shuffle=True,
        random_state=1,
    )

    preprocessor = build_common_preprocessor(X)

    # Python 대표 구현
    # J48과 sklearn DecisionTreeClassifier는 동일 알고리즘이 아님.
    models = {
        "ZeroR": DummyClassifier(strategy="most_frequent"),
        "Logistic": LogisticRegression(
            max_iter=5000,
            solver="lbfgs",
        ),
        "DecisionTree(CART)": DecisionTreeClassifier(
            random_state=1,
        ),
        "RandomForest": RandomForestClassifier(
            n_estimators=100,
            random_state=1,
            n_jobs=-1,
        ),
    }

    rows = []
    confusion_rows = []

    for name, model in models.items():
        pipe = Pipeline(
            steps=[
                ("preprocess", preprocessor),
                ("model", model),
            ]
        )

        # 각 샘플은 정확히 한 번씩 test fold의 예측값을 가짐
        pred = cross_val_predict(
            pipe,
            X,
            y,
            cv=cv,
            method="predict",
            n_jobs=1,
        )

        # ROC-AUC 계산용 확률
        proba = cross_val_predict(
            pipe,
            X,
            y,
            cv=cv,
            method="predict_proba",
            n_jobs=1,
        )[:, 1]

        tn, fp, fn, tp = confusion_matrix(
            y,
            pred,
            labels=[0, 1],
        ).ravel()

        row = {
            "Model": name,
            "Accuracy": accuracy_score(y, pred),
            "Kappa": cohen_kappa_score(y, pred),
            "ROC_AUC": roc_auc_score(y, proba),
            "Risk_Recall": recall_score(y, pred, pos_label=1),
            "Risk_Precision": precision_score(
                y, pred, pos_label=1, zero_division=0
            ),
            "Risk_F1": f1_score(
                y, pred, pos_label=1, zero_division=0
            ),
            "TN": int(tn),
            "FP": int(fp),
            "FN": int(fn),
            "TP": int(tp),
        }
        rows.append(row)

        confusion_rows.extend([
            {
                "Model": name,
                "Actual": "nonrisk",
                "Predicted_nonrisk": int(tn),
                "Predicted_risk": int(fp),
            },
            {
                "Model": name,
                "Actual": "risk",
                "Predicted_nonrisk": int(fn),
                "Predicted_risk": int(tp),
            },
        ])

    metrics = pd.DataFrame(rows)
    save_csv(metrics, output_dir / "03_python_10fold_metrics.csv")

    confusion_df = pd.DataFrame(confusion_rows)
    save_csv(confusion_df, output_dir / "04_confusion_matrices.csv")

    return metrics


# ---------------------------------------------------------------------
# 5. Weka vs Python 비교표
# ---------------------------------------------------------------------

def make_weka_python_comparison(python_metrics, output_dir):
    py = python_metrics.set_index("Model")

    mapping = [
        ("ZeroR", "ZeroR"),
        ("Logistic", "Logistic"),
        ("J48", "DecisionTree(CART)"),
        ("RandomForest", "RandomForest"),
    ]

    rows = []

    for weka_name, py_name in mapping:
        w = WEKA_RESULTS[weka_name]
        p = py.loc[py_name]

        rows.append({
            "Weka_Model": weka_name,
            "Python_Model": py_name,
            "Weka_Accuracy": w["Accuracy"],
            "Python_Accuracy": p["Accuracy"],
            "Accuracy_Diff_pp": (p["Accuracy"] - w["Accuracy"]) * 100,
            "Weka_Kappa": w["Kappa"],
            "Python_Kappa": p["Kappa"],
            "Weka_ROC_AUC": w["ROC_AUC"],
            "Python_ROC_AUC": p["ROC_AUC"],
            "Weka_Risk_Recall": w["Risk_Recall"],
            "Python_Risk_Recall": p["Risk_Recall"],
        })

    comp = pd.DataFrame(rows)
    save_csv(comp, output_dir / "05_weka_python_comparison.csv")

    # Accuracy 비교 차트
    labels = comp["Weka_Model"].tolist()
    x = np.arange(len(labels))
    width = 0.36

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(x - width / 2, comp["Weka_Accuracy"] * 100, width, label="Weka")
    ax.bar(x + width / 2, comp["Python_Accuracy"] * 100, width, label="Python")
    ax.set_title("Weka vs Python accuracy")
    ax.set_xlabel("Model")
    ax.set_ylabel("Accuracy (%)")
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 100)
    ax.legend()
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_dir / "chart_weka_vs_python_accuracy.png", dpi=180)
    plt.close(fig)

    # Risk Recall 비교 차트
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(x - width / 2, comp["Weka_Risk_Recall"] * 100, width, label="Weka")
    ax.bar(x + width / 2, comp["Python_Risk_Recall"] * 100, width, label="Python")
    ax.set_title("Weka vs Python risk recall")
    ax.set_xlabel("Model")
    ax.set_ylabel("Risk recall (%)")
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 100)
    ax.legend()
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_dir / "chart_weka_vs_python_risk_recall.png", dpi=180)
    plt.close(fig)

    return comp


# ---------------------------------------------------------------------
# 6. Logistic 오즈비(OR)
# ---------------------------------------------------------------------

def fit_logistic_odds_ratios(df, output_dir):
    """
    해석을 위해 기준범주를 명확하게 설정하고 전체 데이터에 Logistic을 적합.
    교차검증 성능표와 목적이 다름:
    - CV: 일반화 성능 평가
    - 여기: 변수 방향/상대적 영향 해석
    """
    X = df[EXPECTED_FEATURES].copy()
    y = (df[TARGET] == "risk").astype(int)

    categorical_cols = [c for c in EXPECTED_FEATURES if c not in NUMERIC_COLS]

    # 실제 ARFF 값과 지정 범주가 맞는지 확인
    categories = []
    for c in categorical_cols:
        expected = CATEGORY_ORDERS[c]
        actual = set(X[c].dropna().unique())
        missing_from_order = actual - set(expected)

        if missing_from_order:
            raise ValueError(
                f"{c}에 CATEGORY_ORDERS에 없는 값이 있습니다: "
                f"{sorted(missing_from_order)}"
            )
        categories.append(expected)

    interpret_preprocessor = ColumnTransformer(
        transformers=[
            (
                "cat",
                make_ohe(
                    drop="first",
                    categories=categories,
                ),
                categorical_cols,
            ),
            (
                "num",
                "passthrough",
                NUMERIC_COLS,
            ),
        ]
    )

    model = LogisticRegression(
        max_iter=5000,
        solver="lbfgs",
    )

    pipe = Pipeline([
        ("preprocess", interpret_preprocessor),
        ("model", model),
    ])

    pipe.fit(X, y)

    feature_names = pipe.named_steps["preprocess"].get_feature_names_out()
    coef = pipe.named_steps["model"].coef_[0]

    or_df = pd.DataFrame({
        "Feature": feature_names,
        "Coefficient": coef,
        "Odds_Ratio": np.exp(coef),
    })

    # 보기 편한 이름
    or_df["Feature"] = (
        or_df["Feature"]
        .str.replace("cat__", "", regex=False)
        .str.replace("num__", "", regex=False)
    )

    or_df["Direction"] = np.where(
        or_df["Odds_Ratio"] > 1,
        "risk 증가 방향",
        "risk 감소 방향",
    )

    or_df["Abs_Log_OR"] = np.abs(np.log(or_df["Odds_Ratio"]))
    or_df = or_df.sort_values("Abs_Log_OR", ascending=False)

    save_csv(
        or_df.drop(columns=["Abs_Log_OR"]),
        output_dir / "06_logistic_odds_ratios.csv",
    )

    # 기준범주 표
    ref_rows = []
    for c in categorical_cols:
        ref_rows.append({
            "Variable": c,
            "Reference_Category": CATEGORY_ORDERS[c][0],
        })

    ref_df = pd.DataFrame(ref_rows)
    save_csv(ref_df, output_dir / "07_logistic_reference_categories.csv")

    return or_df


# ---------------------------------------------------------------------
# 7. RandomForest 변수 중요도
# ---------------------------------------------------------------------

def original_feature_from_encoded(encoded_name):
    """One-hot feature name을 원래 변수명으로 묶기."""
    name = encoded_name.replace("cat__", "").replace("num__", "")

    for original in EXPECTED_FEATURES:
        if name == original or name.startswith(original + "_"):
            return original
    return name


def fit_random_forest_importance(df, output_dir):
    X = df[EXPECTED_FEATURES].copy()
    y = (df[TARGET] == "risk").astype(int)

    preprocessor = build_common_preprocessor(X)

    model = RandomForestClassifier(
        n_estimators=100,
        random_state=1,
        n_jobs=-1,
    )

    pipe = Pipeline([
        ("preprocess", preprocessor),
        ("model", model),
    ])

    pipe.fit(X, y)

    feature_names = pipe.named_steps["preprocess"].get_feature_names_out()
    importance = pipe.named_steps["model"].feature_importances_

    encoded_df = pd.DataFrame({
        "Encoded_Feature": feature_names,
        "Importance": importance,
    }).sort_values("Importance", ascending=False)

    encoded_df["Original_Feature"] = encoded_df["Encoded_Feature"].apply(
        original_feature_from_encoded
    )

    save_csv(
        encoded_df,
        output_dir / "08_randomforest_importance_encoded.csv",
    )

    grouped = (
        encoded_df.groupby("Original_Feature", as_index=False)["Importance"]
        .sum()
        .sort_values("Importance", ascending=False)
    )

    save_csv(
        grouped,
        output_dir / "09_randomforest_importance_grouped.csv",
    )

    fig, ax = plt.subplots(figsize=(9, 6))
    plot_df = grouped.sort_values("Importance", ascending=True)
    ax.barh(plot_df["Original_Feature"], plot_df["Importance"])
    ax.set_title("RandomForest grouped feature importance")
    ax.set_xlabel("Importance")
    ax.set_ylabel("Feature")
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(
        output_dir / "chart_randomforest_feature_importance.png",
        dpi=180,
    )
    plt.close(fig)

    return encoded_df, grouped


# ---------------------------------------------------------------------
# 8. 발표용 자동 요약
# ---------------------------------------------------------------------

def make_presentation_summary(
    df,
    python_metrics,
    comparison,
    odds_ratios,
    rf_grouped,
    output_dir,
):
    total = len(df)
    risk_n = int((df[TARGET] == "risk").sum())
    nonrisk_n = int((df[TARGET] == "nonrisk").sum())
    risk_rate = risk_n / total

    py_idx = python_metrics.set_index("Model")

    # Logistic 주요 OR 중 지역 변수는 우선 제외하여 발표 해석을 단순화
    or_clean = odds_ratios[
        ~odds_ratios["Feature"].str.startswith("region_")
    ].copy()

    top_risk = or_clean[or_clean["Odds_Ratio"] > 1].head(5)
    top_protective = (
        or_clean[or_clean["Odds_Ratio"] < 1]
        .sort_values("Odds_Ratio")
        .head(5)
    )

    rf_top = rf_grouped.head(5)

    lines = []
    lines.append("고령자 디지털 소외 위험 예측 - Python 분석 요약")
    lines.append("=" * 60)
    lines.append("")
    lines.append("[1] 데이터")
    lines.append(f"- 전체: {total:,}명")
    lines.append(f"- risk: {risk_n:,}명 ({risk_rate*100:.1f}%)")
    lines.append(f"- nonrisk: {nonrisk_n:,}명 ({(1-risk_rate)*100:.1f}%)")
    lines.append(f"- ZeroR 기준선: {risk_rate*100:.1f}%")
    lines.append("")
    lines.append("[2] Python 10-fold 교차검증 결과")

    for model_name in [
        "ZeroR",
        "Logistic",
        "DecisionTree(CART)",
        "RandomForest",
    ]:
        r = py_idx.loc[model_name]
        lines.append(
            f"- {model_name}: "
            f"Accuracy {r['Accuracy']*100:.2f}%, "
            f"Kappa {r['Kappa']:.3f}, "
            f"ROC-AUC {r['ROC_AUC']:.3f}, "
            f"risk Recall {r['Risk_Recall']*100:.2f}%"
        )

    lines.append("")
    lines.append("[3] Weka와 Python 비교")
    for _, r in comparison.iterrows():
        lines.append(
            f"- {r['Weka_Model']} vs {r['Python_Model']}: "
            f"Weka Accuracy {r['Weka_Accuracy']*100:.1f}% / "
            f"Python {r['Python_Accuracy']*100:.2f}% "
            f"(차이 {r['Accuracy_Diff_pp']:+.2f}%p)"
        )

    lines.append("")
    lines.append(
        "- Logistic과 RandomForest는 두 환경에서 성능 경향이 유사한지 확인."
    )
    lines.append(
        "- J48과 Python DecisionTree는 서로 다른 트리 알고리즘(C4.5 vs CART)이므로 "
        "성능 차이를 프로그램 우열로 단정하지 않음."
    )

    lines.append("")
    lines.append("[4] Logistic 영향 방향 - 참고용")
    for _, r in top_risk.iterrows():
        lines.append(
            f"- risk 증가 방향: {r['Feature']} "
            f"(OR {r['Odds_Ratio']:.2f})"
        )
    for _, r in top_protective.iterrows():
        lines.append(
            f"- risk 감소 방향: {r['Feature']} "
            f"(OR {r['Odds_Ratio']:.2f})"
        )

    lines.append("")
    lines.append("[5] RandomForest 주요 변수")
    for _, r in rf_top.iterrows():
        lines.append(
            f"- {r['Original_Feature']}: importance {r['Importance']:.4f}"
        )

    lines.append("")
    lines.append("[6] 해석 주의")
    lines.append(
        "- Weka와 Python의 10-fold 내부 분할은 완전히 동일하다고 보장하지 않음."
    )
    lines.append(
        "- 모델 구현 및 기본 파라미터 차이 때문에 일부 성능 차이는 자연스럽게 발생할 수 있음."
    )
    lines.append(
        "- 정보화교육 경험자는 표본이 매우 적으므로 관련 효과는 신중하게 해석."
    )
    lines.append(
        "- 분석 결과는 인과관계가 아니라 예측 및 연관성 중심으로 해석."
    )

    summary_path = output_dir / "10_presentation_summary.txt"
    summary_path.write_text("\n".join(lines), encoding="utf-8")

    return lines


# ---------------------------------------------------------------------
# 9. 메인 실행
# ---------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="고령자 디지털 소외 위험 예측 - Weka vs Python 비교분석"
    )
    parser.add_argument(
        "--input",
        type=str,
        default=None,
        help="ARFF 입력 파일 경로",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="결과 저장 폴더 경로",
    )
    args = parser.parse_args()

    # 입력파일 결정
    if args.input:
        input_path = Path(args.input).expanduser().resolve()
    else:
        found = find_default_input()
        if found is None:
            print(
                f"[오류] '{DEFAULT_FILENAME}' 파일을 찾을 수 없습니다.\n"
                "ARFF 파일과 이 Python 파일을 같은 폴더에 두거나,\n"
                "--input 옵션으로 경로를 지정하세요."
            )
            sys.exit(1)
        input_path = found.resolve()

    if not input_path.exists():
        print(f"[오류] 입력 파일이 존재하지 않습니다: {input_path}")
        sys.exit(1)

    # 출력폴더 결정
    if args.output:
        output_dir = Path(args.output).expanduser().resolve()
    else:
        output_dir = input_path.parent / "python_analysis_outputs"

    output_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 72)
    print("고령자 디지털 소외 위험 예측 - Weka vs Python 비교분석")
    print("=" * 72)
    print(f"입력파일: {input_path}")
    print(f"결과폴더: {output_dir}")
    print()

    # 1) 데이터
    print("[1/7] 데이터 로드 및 검증")
    df = load_and_validate(input_path)

    risk_n = int((df[TARGET] == "risk").sum())
    nonrisk_n = int((df[TARGET] == "nonrisk").sum())

    print(f"  - 전체: {len(df):,}명")
    print(f"  - 입력변수: {len(EXPECTED_FEATURES)}개")
    print(f"  - risk: {risk_n:,}명 ({risk_n/len(df)*100:.1f}%)")
    print(f"  - nonrisk: {nonrisk_n:,}명 ({nonrisk_n/len(df)*100:.1f}%)")

    # 2) 기술통계
    print("[2/7] 기술통계 및 그룹별 소외율")
    class_dist, group_rates = make_descriptive_outputs(df, output_dir)

    # 3) CV
    print("[3/7] Python 10-fold 교차검증")
    python_metrics = evaluate_models(df, output_dir)

    display_metrics = python_metrics[
        [
            "Model",
            "Accuracy",
            "Kappa",
            "ROC_AUC",
            "Risk_Recall",
            "Risk_Precision",
            "Risk_F1",
        ]
    ].copy()

    for c in [
        "Accuracy",
        "Risk_Recall",
        "Risk_Precision",
        "Risk_F1",
    ]:
        display_metrics[c] = (display_metrics[c] * 100).round(2)

    display_metrics["Kappa"] = display_metrics["Kappa"].round(3)
    display_metrics["ROC_AUC"] = display_metrics["ROC_AUC"].round(3)

    print()
    print(display_metrics.to_string(index=False))
    print()

    # 4) Weka 비교
    print("[4/7] Weka vs Python 비교표")
    comparison = make_weka_python_comparison(
        python_metrics,
        output_dir,
    )

    # 5) Logistic OR
    print("[5/7] Logistic 오즈비(OR)")
    odds_ratios = fit_logistic_odds_ratios(df, output_dir)

    # 6) RF importance
    print("[6/7] RandomForest 변수 중요도")
    rf_encoded, rf_grouped = fit_random_forest_importance(
        df,
        output_dir,
    )

    # 7) 발표 요약
    print("[7/7] 발표용 요약 생성")
    summary_lines = make_presentation_summary(
        df,
        python_metrics,
        comparison,
        odds_ratios,
        rf_grouped,
        output_dir,
    )

    print()
    print("=" * 72)
    print("분석 완료")
    print("=" * 72)
    print("\n".join(summary_lines[:25]))
    print()
    print(f"전체 결과는 다음 폴더에 저장되었습니다:\n{output_dir}")
    print()
    print("주요 결과 파일:")
    print("  01_class_distribution.csv")
    print("  02_group_risk_rates.csv")
    print("  03_python_10fold_metrics.csv")
    print("  04_confusion_matrices.csv")
    print("  05_weka_python_comparison.csv")
    print("  06_logistic_odds_ratios.csv")
    print("  07_logistic_reference_categories.csv")
    print("  08_randomforest_importance_encoded.csv")
    print("  09_randomforest_importance_grouped.csv")
    print("  10_presentation_summary.txt")
    print("  chart_*.png")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("\n" + "=" * 72)
        print("실행 중 오류가 발생했습니다.")
        print("=" * 72)
        print(f"{type(e).__name__}: {e}")
        print("\n상세 오류:")
        traceback.print_exc()
    finally:
        try:
            input("\nEnter 키를 누르면 창이 닫힙니다...")
        except EOFError:
            pass
