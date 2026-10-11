# SoilClass image classifier

This project trains a five-class soil image classifier using MobileNetV3 transfer learning. It includes training, test evaluation, single-image prediction, and a small browser interface.

## Dataset

The expected dataset structure is:

```text
sample/
  train_data/type_1 ... type_5
  valid_data/type_1 ... type_5
  test_data/type_1  ... type_5
```

The supplied dataset contains 552 images: 389 training, 55 validation, and 108 test images.

## Setup on Windows

Install Python 3.11 or 3.12, then open PowerShell in this folder:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Train and evaluate

From this project folder, run:

```powershell
python train.py --data "..\SoilClass A Novel Image Dataset for Soil Classific\SoilClass A Novel Image Dataset for Soil Classific\sample"
```

The best model and evaluation artifacts are written to `outputs/`:

- `best_model.pt`: trained model
- `metadata.json`: classes and headline accuracy
- `test_metrics.csv`: precision, recall, and F1-score per class
- `training_history.png`: training curves
- `confusion_matrix.png`: test confusion matrix

Training uses a GPU automatically when CUDA is available. On CPU it will take longer.

## Predict one image

```powershell
python predict.py "path\to\soil-image.jpg"
```

## Start the browser interface

```powershell
python app.py
```

Open the local address printed in the terminal, usually `http://127.0.0.1:7860`.

## Important limitation

The publisher labels the classes only as `type_1` through `type_5`; it does not provide physical soil names. The model therefore predicts those dataset labels, not names such as sandy or clay soil.
