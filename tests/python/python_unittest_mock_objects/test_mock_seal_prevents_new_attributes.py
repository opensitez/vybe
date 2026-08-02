# vybe-test: python/python_unittest_mock_objects/test_mock_seal_prevents_new_attributes
# origin: languages/python/tests/python/test_python_unittest_mock_objects.rs

from unittest.mock import Mock, seal, sys
if sys.version_info >= (3, 8):
    m = Mock()
    m.existing = 1
    seal(m)
    try:
        m.new_attr = 2
    except AttributeError:
        print("AttributeError")
else:
    print("AttributeError")
