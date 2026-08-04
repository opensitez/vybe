// vybe-test: dart/enums_advanced/enum_as_class_field
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Role { admin, user, guest }
class User {
  String name;
  Role role;
  User(this.name, this.role);
}

void main() {}
