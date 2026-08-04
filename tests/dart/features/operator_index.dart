// vybe-test: dart/features/operator_index
// origin: languages/dart/tests/dart/test_features.rs

class Grid { List data; Grid(this.data); int operator [](int i) { return data[i]; } }

void main() {}
