// vybe-test: dart/dart_apis/mixins
// origin: languages/dart/tests/dart/test_dart_apis.rs

mixin Greetable { String greet() => 'Hello'; } class Person with Greetable { String name; Person(this.name); }

void main() {}
