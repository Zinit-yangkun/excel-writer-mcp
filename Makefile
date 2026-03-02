test:
	uv run pytest

test-cov:
	uv run pytest --cov --cov-report=term-missing

clean:
	rm -rf dist/

build: clean
	uv build

publish: build
	uv publish

run:
	uv run excel-writer-mcp

.PHONY: test test-cov clean build publish run
