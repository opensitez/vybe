// vybe-test: dart/const_final/late_initialized_field
// origin: languages/dart/tests/dart/test_const_final.rs

class Db { late String connection; Db() { connection = 'sqlite:memory'; } String get conn => connection; }

void main() {}
