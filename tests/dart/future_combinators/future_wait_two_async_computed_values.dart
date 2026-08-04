// vybe-test: dart/future_combinators/future_wait_two_async_computed_values
// origin: languages/dart/tests/dart/test_future_combinators.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

Future<int> square(int n) async => n * n;
Future<void> __vybeMain() async {
  var results = await Future.wait([square(2), square(3)]);
  __p(results.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('4,9');
}
