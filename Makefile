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

.PHONY: test test-cov clean build publish
