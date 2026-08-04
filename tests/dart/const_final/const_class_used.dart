// vybe-test: dart/const_final/const_class_used
// origin: languages/dart/tests/dart/test_const_final.rs

class Size { final int w; final int h; const Size(this.w, this.h); } void main() { const s = Size(100, 200); print(s.w); }