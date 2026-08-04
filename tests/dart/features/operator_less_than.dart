// vybe-test: dart/features/operator_less_than
// origin: languages/dart/tests/dart/test_features.rs

class Score { int v; Score(this.v); bool operator <(Score other) { return v < other.v; } }

void main() {}
