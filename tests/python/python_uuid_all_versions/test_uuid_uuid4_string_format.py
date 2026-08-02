# vybe-test: python/python_uuid_all_versions/test_uuid_uuid4_string_format
# origin: languages/python/tests/python/test_python_uuid_all_versions.rs

import uuid
u = uuid.uuid4()
s = str(u)
parts = s.split("-")
print(len(parts))
print([len(p) for p in parts])
