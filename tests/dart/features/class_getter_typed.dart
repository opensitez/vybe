// vybe-test: dart/features/class_getter_typed
// origin: languages/dart/tests/dart/test_features.rs

class Rect { int w; int h; Rect(this.w, this.h); int get area { return w * h; } }

void main() {}
