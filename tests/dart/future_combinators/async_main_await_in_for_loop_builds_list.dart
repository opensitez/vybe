// vybe-test: dart/future_combinators/async_main_await_in_for_loop_builds_list
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

Future<int> id(int n) async => n;
Future<void> __vybeMain() async {
  var out = <int>[];
  for (var i = 1; i <= 3; i++) {
    out.add(await id(i));
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3');
}
