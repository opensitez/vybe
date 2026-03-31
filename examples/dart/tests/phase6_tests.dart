class Point {
  double x, y;
  Point(this.x, this.y);
  
  // Redirecting constructor
  Point.zero() : this(0, 0);
  
  void move(double dx, double dy) {
    x += dx;
    y += dy;
  }
}

void testCascades() {
  print("--- Testing Cascades ---");
  var p = Point(1, 2);
  p..move(10, 10)..x = 100;
  print("Point x: ${p.x}, y: ${p.y}"); // Should be 100, 12
}

void testRedirectingConstructors() {
  print("--- Testing Redirecting Constructors ---");
  var p = Point.zero();
  print("Point.zero x: ${p.x}, y: ${p.y}");
}

void testTryCatchMulti() {
  print("--- Testing Multi-Catch ---");
  try {
    throw "Error String";
  } on int catch (e) {
    print("Caught int: $e");
  } on String catch (e) {
    print("Caught String: $e");
  } finally {
    print("Finally reached");
  }
}

typedef NumberList = List;

void testTypedefs() {
  print("--- Testing Typedefs ---");
  NumberList list = [1, 2, 3];
  print("List length: ${list.length}");
}

void main() {
  testCascades();
  testRedirectingConstructors();
  testTryCatchMulti();
  testTypedefs();
}
