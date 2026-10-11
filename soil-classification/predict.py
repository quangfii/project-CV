from __future__ import annotations

import argparse
from pathlib import Path

import torch
from PIL import Image

from soil_classifier import load_trained_model, predict_image


def main():
    parser = argparse.ArgumentParser(description="Predict the soil class of one image.")
    parser.add_argument("image", type=Path)
    parser.add_argument("--model", type=Path, default=Path("outputs/best_model.pt"))
    args = parser.parse_args()
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model, class_names = load_trained_model(args.model, device)
    with Image.open(args.image) as image:
        probabilities = predict_image(model, image, class_names, device)
    for class_name, probability in sorted(probabilities.items(), key=lambda item: item[1], reverse=True):
        print(f"{class_name}: {probability:.2%}")


if __name__ == "__main__":
    main()
