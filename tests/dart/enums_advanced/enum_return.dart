// vybe-test: dart/enums_advanced/enum_return
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Result { ok, err }
Result check(int x) { return x > 0 ? Result.ok : Result.err; }

void main() {}
