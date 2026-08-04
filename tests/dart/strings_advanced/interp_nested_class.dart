// vybe-test: dart/strings_advanced/interp_nested_class
// origin: languages/dart/tests/dart/test_strings_advanced.rs

class Dog { String name; Dog(this.name); } void main() { var d = Dog('Rex'); var s = 'Dog: ${d.name}'; }