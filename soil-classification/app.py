from __future__ import annotations

from pathlib import Path

import gradio as gr
import torch

from soil_classifier import load_trained_model, predict_image

MODEL_PATH = Path(__file__).parent / "outputs" / "best_model.pt"
DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
MODEL, CLASS_NAMES = load_trained_model(MODEL_PATH, DEVICE)


def classify(image):
    if image is None:
        return {}
    return predict_image(MODEL, image, CLASS_NAMES, DEVICE)


demo = gr.Interface(
    fn=classify,
    inputs=gr.Image(type="pil", label="Soil image"),
    outputs=gr.Label(num_top_classes=5, label="Prediction"),
    title="SoilClass",
    description="Upload a clear soil photo to classify it as type_1 through type_5.",
)

if __name__ == "__main__":
    demo.launch()
