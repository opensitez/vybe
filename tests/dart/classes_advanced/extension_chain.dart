// vybe-test: dart/classes_advanced/extension_chain
// origin: languages/dart/tests/dart/test_classes_advanced.rs

extension StringX on String {
  String shout() => toUpperCase() + '!';
}
void main() { print('hello'.shout()); }