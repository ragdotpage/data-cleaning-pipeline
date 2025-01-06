# Project Setup and Usage

## Environment Setup

1. Activate the virtual environment:
```bash
source .venv/bin/activate
```

## Installing Dependencies

You can install dependencies using either `pip` or `uv` (recommended):

```bash
# Using uv (recommended - faster installation)
uv pip install requirements.txt

# Or using traditional pip
pip install requirements.txt
```

> **Note**: We recommend using `uv` as it provides significantly faster package installation and dependency resolution.

## Running the Application

You can run the application using either Python directly or `uv`:

```bash
# Using uv
uv run filename.py

# Or using Python directly
python filename.py
```
