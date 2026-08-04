// vybe-test: dart/async_generators/future_when_complete
// origin: languages/dart/tests/dart/test_async_generators.rs

Future.value(1).whenComplete(() => print('done'));

void main() {}
