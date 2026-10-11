from __future__ import annotations

import argparse
import csv
import math
from collections import Counter, defaultdict
from pathlib import Path

import torch
from PIL import Image, UnidentifiedImageError

from soil_classifier import load_trained_model, predict_image

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Evaluate a trained model on an external labeled image collection without retraining."
    )
    parser.add_argument("--data", type=Path, required=True)
    parser.add_argument("--model", type=Path, default=Path("outputs/best_model.pt"))
    parser.add_argument("--output", type=Path, default=Path("outputs/external_test"))
    parser.add_argument("--confidence-threshold", type=float, default=0.60)
    return parser.parse_args()


def normalized_entropy(probabilities: list[float]) -> float:
    entropy = -sum(value * math.log(value) for value in probabilities if value > 0)
    return entropy / math.log(len(probabilities))


def main():
    args = parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model, model_classes = load_trained_model(args.model, device)

    image_paths = sorted(
        path
        for path in args.data.rglob("*")
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    )
    if not image_paths:
        raise ValueError(f"No supported images found under {args.data}")

    rows = []
    failures = []
    for index, path in enumerate(image_paths, start=1):
        external_class = path.relative_to(args.data).parts[0]
        try:
            with Image.open(path) as image:
                scores = predict_image(model, image, model_classes, device)
        except (OSError, UnidentifiedImageError) as error:
            failures.append((str(path), str(error)))
            continue
        predicted_class, confidence = max(scores.items(), key=lambda item: item[1])
        rows.append(
            {
                "path": str(path),
                "external_label": external_class,
                "predicted_label": predicted_class,
                "confidence": confidence,
                "normalized_entropy": normalized_entropy(list(scores.values())),
                **{f"probability_{name}": scores[name] for name in model_classes},
            }
        )
        if index % 100 == 0 or index == len(image_paths):
            print(f"Processed {index}/{len(image_paths)} images")

    prediction_fields = [
        "path",
        "external_label",
        "predicted_label",
        "confidence",
        "normalized_entropy",
        *[f"probability_{name}" for name in model_classes],
    ]
    with (args.output / "predictions.csv").open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=prediction_fields)
        writer.writeheader()
        writer.writerows(rows)

    grouped = defaultdict(list)
    for row in rows:
        grouped[row["external_label"]].append(row)

    summary_rows = []
    for external_label, items in sorted(grouped.items()):
        predictions = Counter(item["predicted_label"] for item in items)
        dominant_label, dominant_count = predictions.most_common(1)[0]
        summary_rows.append(
            {
                "external_label": external_label,
                "images": len(items),
                "dominant_prediction": dominant_label,
                "dominant_share": dominant_count / len(items),
                "mean_confidence": sum(item["confidence"] for item in items) / len(items),
                "low_confidence_share": sum(
                    item["confidence"] < args.confidence_threshold for item in items
                )
                / len(items),
                "mean_normalized_entropy": sum(
                    item["normalized_entropy"] for item in items
                )
                / len(items),
                **{f"predicted_{name}": predictions[name] for name in model_classes},
            }
        )

    summary_fields = [
        "external_label",
        "images",
        "dominant_prediction",
        "dominant_share",
        "mean_confidence",
        "low_confidence_share",
        "mean_normalized_entropy",
        *[f"predicted_{name}" for name in model_classes],
    ]
    with (args.output / "summary.csv").open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=summary_fields)
        writer.writeheader()
        writer.writerows(summary_rows)

    if failures:
        with (args.output / "failures.csv").open("w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["path", "error"])
            writer.writerows(failures)

    print(f"Predictions: {len(rows)} | Failures: {len(failures)}")
    print(f"Reports saved to: {args.output.resolve()}")


if __name__ == "__main__":
    main()
