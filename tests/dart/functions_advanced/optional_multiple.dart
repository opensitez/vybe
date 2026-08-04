// vybe-test: dart/functions_advanced/optional_multiple
// origin: languages/dart/tests/dart/test_functions_advanced.rs

String format(String s, [int width = 10, String fill = ' ']) { return s.padLeft(width, fill); }

void main() {}
