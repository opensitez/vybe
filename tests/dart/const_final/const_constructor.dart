// vybe-test: dart/const_final/const_constructor
// origin: languages/dart/tests/dart/test_const_final.rs

class Color { final int r; final int g; final int b; const Color(this.r, this.g, this.b); } const red = Color(255, 0, 0);

void main() {}
