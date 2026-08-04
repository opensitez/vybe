// vybe-test: dart/await_for_loops/await_for_sequential_two_generators
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

Stream<int> first() async* { yield 1; yield 2; }
Stream<int> second() async* { yield 3; yield 4; }
Future<void> __vybeMain() async {
  var out = <int>[];
  await for (var v in first()) { out.add(v); }
  await for (var v in second()) { out.add(v); }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3,4');
}
