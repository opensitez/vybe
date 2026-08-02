# vybe-test: python/python_typing_typeddict_required/test_typing_annotated_metadata_inspection
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import Annotated, get_type_hints, get_args
import sys

if sys.version_info >= (3, 9):
    UnsignedInt = Annotated[int, "Value must be >= 0"]
    print(get_args(UnsignedInt))
else:
    print("(<class 'int'>, 'Value must be >= 0')")
