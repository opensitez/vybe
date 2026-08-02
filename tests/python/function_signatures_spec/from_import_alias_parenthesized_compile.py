# vybe-test: python/function_signatures_spec/from_import_alias_parenthesized_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

from package.module import (
    name as alias,
    other
)
