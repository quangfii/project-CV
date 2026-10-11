from __future__ import annotations

import argparse
import copy
import csv
import random
import time
from pathlib import Path

import matplotlib.pyplot as plt
import torch
from torch import nn
from torch.utils.data import DataLoader
from torchvision import datasets, transforms

from soil_classifier import IMAGE_SIZE, MEAN, STD, build_model, save_metadata


def parse_args():
    parser = argparse.ArgumentParser(description="Train a five-class soil image classifier.")
    parser.add_argument("--data", type=Path, required=True, help="Folder containing train_data, valid_data and test_data")
    parser.add_argument("--output", type=Path, default=Path("outputs"))
    parser.add_argument("--epochs", type=int, default=20)
    parser.add_argument("--batch-size", type=int, default=16)
    parser.add_argument("--learning-rate", type=float, default=3e-4)
    parser.add_argument("--workers", type=int, default=0)
    parser.add_argument("--seed", type=int, default=42)
    return parser.parse_args()


def seed_everything(seed: int) -> None:
    random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def make_loaders(data_dir: Path, batch_size: int, workers: int):
    train_transform = transforms.Compose(
        [
            transforms.RandomResizedCrop(IMAGE_SIZE, scale=(0.7, 1.0)),
            transforms.RandomHorizontalFlip(),
            transforms.RandomRotation(12),
            transforms.ColorJitter(brightness=0.2, contrast=0.2, saturation=0.15),
            transforms.ToTensor(),
            transforms.Normalize(MEAN, STD),
        ]
    )
    eval_transform = transforms.Compose(
        [
            transforms.Resize(256),
            transforms.CenterCrop(IMAGE_SIZE),
            transforms.ToTensor(),
            transforms.Normalize(MEAN, STD),
        ]
    )
    datasets_by_split = {
        "train": datasets.ImageFolder(data_dir / "train_data", train_transform),
        "valid": datasets.ImageFolder(data_dir / "valid_data", eval_transform),
        "test": datasets.ImageFolder(data_dir / "test_data", eval_transform),
    }
    if datasets_by_split["train"].classes != datasets_by_split["valid"].classes or datasets_by_split["train"].classes != datasets_by_split["test"].classes:
        raise ValueError("Class folders differ between train, validation and test splits.")
    loaders = {
        name: DataLoader(
            dataset,
            batch_size=batch_size,
            shuffle=name == "train",
            num_workers=workers,
            pin_memory=torch.cuda.is_available(),
        )
        for name, dataset in datasets_by_split.items()
    }
    return datasets_by_split, loaders


def run_epoch(model, loader, criterion, device, optimizer=None):
    training = optimizer is not None
    model.train(training)
    total_loss = total_correct = total_samples = 0
    for images, labels in loader:
        images, labels = images.to(device), labels.to(device)
        if training:
            optimizer.zero_grad(set_to_none=True)
        with torch.set_grad_enabled(training):
            logits = model(images)
            loss = criterion(logits, labels)
            if training:
                loss.backward()
                optimizer.step()
        total_loss += loss.item() * labels.size(0)
        total_correct += (logits.argmax(1) == labels).sum().item()
        total_samples += labels.size(0)
    return total_loss / total_samples, total_correct / total_samples


@torch.inference_mode()
def evaluate(model, loader, device, num_classes):
    model.eval()
    matrix = torch.zeros((num_classes, num_classes), dtype=torch.int64)
    for images, labels in loader:
        predictions = model(images.to(device)).argmax(1).cpu()
        for actual, predicted in zip(labels, predictions):
            matrix[actual, predicted] += 1
    return matrix


def metrics_from_matrix(matrix: torch.Tensor, class_names: list[str]):
    rows = []
    for index, class_name in enumerate(class_names):
        tp = matrix[index, index].item()
        fp = matrix[:, index].sum().item() - tp
        fn = matrix[index, :].sum().item() - tp
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        rows.append((class_name, precision, recall, f1, int(matrix[index, :].sum())))
    accuracy = matrix.diag().sum().item() / matrix.sum().item()
    return rows, accuracy


def save_plots(history, matrix, class_names, output_dir):
    fig, axes = plt.subplots(1, 2, figsize=(11, 4))
    axes[0].plot(history["train_loss"], label="Train")
    axes[0].plot(history["valid_loss"], label="Validation")
    axes[0].set(title="Loss", xlabel="Epoch")
    axes[0].legend()
    axes[1].plot(history["train_accuracy"], label="Train")
    axes[1].plot(history["valid_accuracy"], label="Validation")
    axes[1].set(title="Accuracy", xlabel="Epoch", ylim=(0, 1))
    axes[1].legend()
    fig.tight_layout()
    fig.savefig(output_dir / "training_history.png", dpi=160)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(6, 5))
    image = ax.imshow(matrix.numpy(), cmap="Blues")
    for row in range(len(class_names)):
        for col in range(len(class_names)):
            ax.text(col, row, int(matrix[row, col]), ha="center", va="center")
    ax.set(xticks=range(len(class_names)), yticks=range(len(class_names)), xticklabels=class_names, yticklabels=class_names, xlabel="Predicted", ylabel="Actual", title="Test confusion matrix")
    fig.colorbar(image, ax=ax)
    fig.tight_layout()
    fig.savefig(output_dir / "confusion_matrix.png", dpi=160)
    plt.close(fig)


def main():
    args = parse_args()
    seed_everything(args.seed)
    args.output.mkdir(parents=True, exist_ok=True)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    datasets_by_split, loaders = make_loaders(args.data, args.batch_size, args.workers)
    class_names = datasets_by_split["train"].classes
    counts = torch.bincount(torch.tensor(datasets_by_split["train"].targets))
    class_weights = counts.sum() / (len(counts) * counts.float())
    criterion = nn.CrossEntropyLoss(weight=class_weights.to(device))

    model = build_model(len(class_names), pretrained=True).to(device)
    for parameter in model.features.parameters():
        parameter.requires_grad = False
    optimizer = torch.optim.AdamW(model.classifier.parameters(), lr=args.learning_rate, weight_decay=1e-4)
    scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode="max", patience=2, factor=0.3)

    history = {key: [] for key in ("train_loss", "train_accuracy", "valid_loss", "valid_accuracy")}
    best_accuracy = -1.0
    best_state = None
    stale_epochs = 0
    print(f"Device: {device} | Classes: {class_names}")
    started = time.time()
    for epoch in range(1, args.epochs + 1):
        train_loss, train_accuracy = run_epoch(model, loaders["train"], criterion, device, optimizer)
        valid_loss, valid_accuracy = run_epoch(model, loaders["valid"], criterion, device)
        for key, value in (("train_loss", train_loss), ("train_accuracy", train_accuracy), ("valid_loss", valid_loss), ("valid_accuracy", valid_accuracy)):
            history[key].append(value)
        scheduler.step(valid_accuracy)
        print(f"Epoch {epoch:02d}/{args.epochs} | train {train_accuracy:.3f} | valid {valid_accuracy:.3f}")
        if valid_accuracy > best_accuracy:
            best_accuracy = valid_accuracy
            best_state = copy.deepcopy(model.state_dict())
            stale_epochs = 0
        else:
            stale_epochs += 1
        if stale_epochs >= 5:
            print("Early stopping: validation accuracy did not improve.")
            break

    model.load_state_dict(best_state)
    checkpoint_path = args.output / "best_model.pt"
    torch.save({"model_state": model.state_dict(), "class_names": class_names, "validation_accuracy": best_accuracy}, checkpoint_path)
    matrix = evaluate(model, loaders["test"], device, len(class_names))
    rows, test_accuracy = metrics_from_matrix(matrix, class_names)
    with (args.output / "test_metrics.csv").open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["class", "precision", "recall", "f1_score", "support"])
        writer.writerows(rows)
        writer.writerow(["overall_accuracy", test_accuracy, "", "", int(matrix.sum())])
    save_plots(history, matrix, class_names, args.output)
    save_metadata(args.output / "metadata.json", class_names, validation_accuracy=best_accuracy, test_accuracy=test_accuracy)
    print(f"Test accuracy: {test_accuracy:.3f}")
    print(f"Saved model and reports to: {args.output.resolve()}")
    print(f"Elapsed: {(time.time() - started) / 60:.1f} minutes")


if __name__ == "__main__":
    main()
