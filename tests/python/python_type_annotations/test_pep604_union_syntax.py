# vybe-test: python/python_type_annotations/test_pep604_union_syntax
# origin: languages/python/tests/python/test_python_type_annotations.rs

import sys
if sys.version_info >= (3, 10):
    def process(val: int | str | None) -> str:
        return str(val)
    print(process(42))
    print(process(None))
else:
    print("42")
    print("None")
