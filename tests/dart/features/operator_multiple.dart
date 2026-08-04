// vybe-test: dart/features/operator_multiple
// origin: languages/dart/tests/dart/test_features.rs

class Num { int v; Num(this.v); Num operator +(Num o) { return Num(v + o.v); } Num operator -(Num o) { return Num(v - o.v); } Num operator *(Num o) { return Num(v * o.v); } }

void main() {}
