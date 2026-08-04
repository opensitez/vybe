// vybe-test: dart/classes_advanced/equality_override
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Point {
  int x; int y;
  Point(this.x, this.y);
  bool operator ==(Object other) {
    if (other is Point) return x == other.x && y == other.y;
    return false;
  }
  int get hashCode => x * 31 + y;
}

void main() {}
