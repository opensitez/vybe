// vybe-test: dart/errors_advanced/throw_string
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() { try { throw 'an error'; } catch (e) { print(e); } }