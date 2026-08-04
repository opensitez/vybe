// vybe-test: dart/await_for_loops/await_for_three_nested_generators_flattened
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

Stream<int> a() async* { yield 1; }
Stream<int> b() async* { yield 2; }
Stream<int> c() async* { yield 3; }
Future<void> __vybeMain() async {
  var out = <int>[];
  await for (var x in a()) {
    await for (var y in b()) {
      await for (var z in c()) {
        out.add(x + y + z);
      }
    }
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('6');
}
