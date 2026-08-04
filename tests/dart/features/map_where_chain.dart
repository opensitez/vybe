// vybe-test: dart/features/map_where_chain
// origin: languages/dart/tests/dart/test_features.rs

void main() { var x = [1,2,3,4,5].map((e) => e * 2).where((e) => e > 4).toList(); }