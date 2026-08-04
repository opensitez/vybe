// vybe-test: dart/async_star_generators/async_star_yield_bool_values
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

Stream<bool> flags() async* { yield true; yield false; }
Future<void> __vybeMain() async {
  var out = <String>[];
  await for (var f in flags()) { out.add('$f'); }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('true,false');
}
