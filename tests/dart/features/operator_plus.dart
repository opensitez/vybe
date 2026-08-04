// vybe-test: dart/features/operator_plus
// origin: languages/dart/tests/dart/test_features.rs

class Vec { int x; int y; Vec(this.x, this.y); Vec operator +(Vec other) { return Vec(x + other.x, y + other.y); } }

void main() {}
