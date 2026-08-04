// vybe-test: dart/collections_advanced/spread_null_aware
// origin: languages/dart/tests/dart/test_collections_advanced.rs

List? a = null; var b = [...?a, 1, 2];

void main() {}
