# Releasing

1. Update `__version__` in `src/htmlexport/__init__.py` and `CHANGELOG.md`.
2. Run the checks: `ruff check src tests && mypy src && pytest`.
3. Commit, push, and wait for CI to pass.
4. Create a GitHub Release with tag `vX.Y.Z`. The `publish.yml` workflow
   builds the package and uploads it to PyPI (Trusted Publishing, no token).

First release only: try the package on TestPyPI first.

    python -m build
    twine upload --repository testpypi dist/*
    pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ "htmlexport[all]"
