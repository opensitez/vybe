// vybe-test: dart/async_generators/future_catch_error
// origin: languages/dart/tests/dart/test_async_generators.rs

Future.error('oops').catchError((e) => print(e));

void main() {}
