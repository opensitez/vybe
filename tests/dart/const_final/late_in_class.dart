// vybe-test: dart/const_final/late_in_class
// origin: languages/dart/tests/dart/test_const_final.rs

class Lazy { late int value; void init() { value = 100; } }

void main() {}
