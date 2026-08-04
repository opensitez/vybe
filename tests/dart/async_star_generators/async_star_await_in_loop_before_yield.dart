// vybe-test: dart/async_star_generators/async_star_await_in_loop_before_yield
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

Future<int> next(int n) async => n + 1;
Stream<int> pipeline(int count) async* {
  var v = 0;
  for (var i = 0; i < count; i++) {
    v = await next(v);
    yield v;
  }
}
Future<void> __vybeMain() async {
  __p(await pipeline(3).join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3');
}
