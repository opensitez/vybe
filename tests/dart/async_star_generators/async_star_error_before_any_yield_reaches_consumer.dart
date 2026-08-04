// vybe-test: dart/async_star_generators/async_star_error_before_any_yield_reaches_consumer
// origin: languages/dart/tests/dart/test_async_star_generators.rs

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

Stream<int> failEarly() async* {
  throw Exception('early');
  yield 1;
}
Future<void> __vybeMain() async {
  var caught = false;
  try {
    await for (var _ in failEarly()) {}
  } catch (_) {
    caught = true;
  }
  __p(caught);
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
