// vybe-test: dart/features/class_named_constructor
// origin: languages/dart/tests/dart/test_features.rs

class Point { int x; int y; Point(this.x, this.y); Point.origin() : this(0, 0); }

void main() {}
