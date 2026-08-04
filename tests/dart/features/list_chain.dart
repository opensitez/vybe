// vybe-test: dart/features/list_chain
// origin: languages/dart/tests/dart/test_features.rs

var x = [1, 2, 3].map((e) => e * 2).where((e) => e > 2).toList();

void main() {}
