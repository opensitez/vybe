// vybe-test: dart/dart_apis/extension_method
// origin: languages/dart/tests/dart/test_dart_apis.rs

extension StringExt on String { String reversed() => split('').reversed.join(''); } var r = 'hello'.reversed();

void main() {}
