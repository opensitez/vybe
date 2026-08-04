// vybe-test: dart/generics/generic_fn_list
// origin: languages/dart/tests/dart/test_generics.rs

List<T> repeat<T>(T val, int times) => List.generate(times, (_) => val);

void main() {}
