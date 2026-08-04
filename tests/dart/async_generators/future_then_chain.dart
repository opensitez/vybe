// vybe-test: dart/async_generators/future_then_chain
// origin: languages/dart/tests/dart/test_async_generators.rs

Future.value(1).then((v) => v * 2).then((v) => print(v));

void main() {}
