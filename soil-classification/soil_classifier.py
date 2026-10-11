from __future__ import annotations

import json
from pathlib import Path

import torch
from PIL import Image
from torch import nn
from torchvision import models, transforms

IMAGE_SIZE = 224
MEAN = (0.485, 0.456, 0.406)
STD = (0.229, 0.224, 0.225)


def build_model(num_classes: int, pretrained: bool = True) -> nn.Module:
    weights = models.MobileNet_V3_Small_Weights.DEFAULT if pretrained else None
    model = models.mobilenet_v3_small(weights=weights)
    model.classifier[3] = nn.Linear(model.classifier[3].in_features, num_classes)
    return model


def inference_transform() -> transforms.Compose:
    return transforms.Compose(
        [
            transforms.Resize(256),
            transforms.CenterCrop(IMAGE_SIZE),
            transforms.ToTensor(),
            transforms.Normalize(MEAN, STD),
        ]
    )


def load_trained_model(checkpoint_path: Path, device: torch.device):
    checkpoint = torch.load(checkpoint_path, map_location=device, weights_only=False)
    class_names = checkpoint["class_names"]
    model = build_model(len(class_names), pretrained=False)
    model.load_state_dict(checkpoint["model_state"])
    model.to(device).eval()
    return model, class_names


@torch.inference_mode()
def predict_image(model, image: Image.Image, class_names: list[str], device):
    tensor = inference_transform()(image.convert("RGB")).unsqueeze(0).to(device)
    probabilities = torch.softmax(model(tensor), dim=1)[0].cpu()
    return {
        class_name: float(probability)
        for class_name, probability in zip(class_names, probabilities)
    }


def save_metadata(path: Path, class_names: list[str], **values) -> None:
    payload = {"class_names": class_names, "image_size": IMAGE_SIZE, **values}
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
