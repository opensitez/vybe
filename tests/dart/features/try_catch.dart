// vybe-test: dart/features/try_catch
// origin: languages/dart/tests/dart/test_features.rs

void main() { try { throw 'error'; } catch (e) { print(e); } }