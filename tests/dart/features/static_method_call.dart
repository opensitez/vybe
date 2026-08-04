// vybe-test: dart/features/static_method_call
// origin: languages/dart/tests/dart/test_features.rs

class Utils { static int double(int x) { return x * 2; } } void main() { var r = Utils.double(5); }