// vybe-test: dart/errors_advanced/exception_in_method
// origin: languages/dart/tests/dart/test_errors_advanced.rs

class Parser {
  int parse(String s) {
    if (s.isEmpty) throw FormatException('empty');
    return int.parse(s);
  }
}

void main() {}
