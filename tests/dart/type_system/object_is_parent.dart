// vybe-test: dart/type_system/object_is_parent
// origin: languages/dart/tests/dart/test_type_system.rs

class Foo {} void main() { var f = Foo(); print(f is Object); }