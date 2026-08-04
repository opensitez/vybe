// vybe-test: dart/async_star_generators/async_star_break_limits_loop_yields
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

Stream<int> capped() async* {
  for (var i = 0; i < 10; i++) {
    if (i == 3) break;
    yield i;
  }
}
Future<void> __vybeMain() async {
  __p(await capped().join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('0,1,2');
}
