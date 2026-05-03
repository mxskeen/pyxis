.PHONY: sync run

sync:
	uv sync

run:
	uv run uvicorn main:app --port 8000
