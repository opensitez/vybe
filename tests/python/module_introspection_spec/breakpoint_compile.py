# vybe-test: python/module_introspection_spec/breakpoint_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs
# `breakpoint()` drops into pdb and HANGS a test run. Neutralising the
# hook keeps the call under test while letting it return.
import sys
sys.breakpointhook = lambda *a, **k: None
breakpoint()
