// vybe-test: dart/classes_advanced/three_level_inheritance
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class A { String name() => 'A'; }
class B extends A { String extra() => 'B'; }
class C extends B { String more() => 'C'; }
void main() { var c = C(); print(c.name()); }