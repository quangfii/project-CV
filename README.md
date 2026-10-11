# Project CV — SoilClass

Computer vision project for five-class soil image classification using MobileNetV3 transfer learning.

The application source, dependencies, training and prediction instructions are in [soil-classification](soil-classification/README.md). Report generation scripts are included at the repository root.

## Run

```powershell
cd soil-classification
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python train.py --data "path\to\sample"
python app.py
```

Supply the SoilClass dataset locally with `train_data`, `valid_data`, and `test_data` folders, each containing `type_1` through `type_5`. Train before starting the app to create `outputs/best_model.pt`.

Datasets, trained models, local environments, generated reports and QA artifacts are excluded from Git. Dataset labels do not establish physical soil types. See the application README for details.
