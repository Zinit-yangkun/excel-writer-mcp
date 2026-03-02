test:
	uv run pytest

test-cov:
	uv run pytest --cov --cov-report=term-missing

.PHONY: test test-cov
