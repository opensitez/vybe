// vybe-test: dart/classes_advanced/private_field
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Account { double _balance = 0; void deposit(double amt) { _balance += amt; } double get balance => _balance; }

void main() {}
