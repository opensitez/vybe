// vybe-test: dart/await_for_loops/await_for_on_async_star_with_await_between_yields
// origin: languages/dart/tests/dart/test_await_for_loops.rs

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

Future<int> bump(int n) async => n + 1;
Stream<int> gen() async* {
  yield await bump(0);
  yield await bump(1);
}
Future<void> __vybeMain() async {
  var out = <int>[];
  await for (var v in gen()) { out.add(v); }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2');
}
