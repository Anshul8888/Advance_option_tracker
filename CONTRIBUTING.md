
```md
# CONTRIBUTING.md

# Contributing

Thanks for your interest in contributing to this project.

## Ways to Contribute

You can help by:

- reporting bugs
- suggesting improvements
- improving documentation
- optimizing performance
- adding new features
- improving error handling and stability

---

## Before You Start

Please:

1. Read the README
2. Understand the current project structure
3. Open an issue first for major changes if possible

---

## Development Guidelines

### 1. Keep changes focused

Try to keep pull requests small and easy to review.

### 2. Preserve project style

Follow the existing code style:

- clear function names
- short docstrings
- modular structure
- descriptive logging

### 3. Avoid breaking config behavior

If you add new configuration options:

- document them in `README.md`
- add sane defaults
- keep backward compatibility where possible

### 4. Be careful with API rate limits

This project interacts with Zerodha APIs, so performance changes should not violate rate-limit assumptions.

### 5. Do not commit secrets

Never commit:

- API keys
- API secrets
- access tokens
- Telegram bot tokens
- personal chat IDs

---

## Suggested Setup

Create a virtual environment and install dependencies:

```bash
python -m venv venv
venv\Scripts\activate
pip install kiteconnect pandas numpy xlwings requests
