// vybe-test: dart/await_for_loops/await_for_string_join_from_events
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

Stream<String> gen() async* { yield 'hello'; yield ' '; yield 'dart'; }
Future<void> __vybeMain() async {
  var buf = '';
  await for (var s in gen()) { buf = buf + s; }
  __p(buf);
}

Future<void> main() async {
  await __vybeMain();
  __check('hello dart');
}
