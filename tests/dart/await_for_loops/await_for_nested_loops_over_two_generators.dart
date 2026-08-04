// vybe-test: dart/await_for_loops/await_for_nested_loops_over_two_generators
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

Stream<int> rows() async* { yield 1; yield 2; }
Stream<int> cols() async* { yield 10; yield 20; }
Future<void> __vybeMain() async {
  var out = <String>[];
  await for (var r in rows()) {
    await for (var c in cols()) {
      out.add('$r$c');
    }
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('110,120,210,220');
}
