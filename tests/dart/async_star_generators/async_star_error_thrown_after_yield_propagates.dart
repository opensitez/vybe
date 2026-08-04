// vybe-test: dart/async_star_generators/async_star_error_thrown_after_yield_propagates
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

Stream<int> bad() async* {
  yield 1;
  throw Exception('boom');
}
Future<void> __vybeMain() async {
  var out = <String>[];
  try {
    await for (var v in bad()) { out.add('$v'); }
  } catch (e) {
    out.add('err');
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,err');
}
