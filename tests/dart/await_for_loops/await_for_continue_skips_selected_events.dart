// vybe-test: dart/await_for_loops/await_for_continue_skips_selected_events
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

Stream<int> gen() async* { for (var i = 1; i <= 5; i++) yield i; }
Future<void> __vybeMain() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v % 2 == 0) continue;
    out.add(v);
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,3,5');
}
