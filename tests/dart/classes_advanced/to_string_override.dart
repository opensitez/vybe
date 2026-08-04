// vybe-test: dart/classes_advanced/to_string_override
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Point { int x; int y; Point(this.x, this.y); String toString() => 'Point($x, $y)'; }

void main() {}
