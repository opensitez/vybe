// vybe-test: dart/features/getter_arrow_body
// origin: languages/dart/tests/dart/test_features.rs

class Circle { double r; Circle(this.r); get area { return 3.14 * r * r; } }

void main() {}
