# vybe-test: python/python_uuid_all_versions/test_uuid_uuid4_uniqueness
# origin: languages/python/tests/python/test_python_uuid_all_versions.rs

import uuid
ids = {str(uuid.uuid4()) for _ in range(100)}
print(len(ids))
