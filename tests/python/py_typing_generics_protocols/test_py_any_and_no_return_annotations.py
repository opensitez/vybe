# vybe-test: python/py_typing_generics_protocols/test_py_any_and_no_return_annotations
# origin: languages/python/tests/python/test_py_typing_generics_protocols.rs

from typing import Any, NoReturn

def log(msg: Any) -> None:
    print(f"LOG: {msg}")

log("test message")
log(123)
