// vybe-test: dart/classes_advanced/redirecting_constructor
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Point { int x; int y; Point(this.x, this.y); Point.origin() : this(0, 0); Point.unit() : this(1, 1); }

void main() {}
