// vybe-test: dart/await_for_loops/await_for_with_local_counter_in_body
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

Stream<int> gen() async* { yield 3; yield 4; yield 5; }
Future<void> __vybeMain() async {
  var idx = 0;
  await for (var v in gen()) {
    __p('$idx:$v');
    idx++;
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('0:3\n1:4\n2:5');
}
