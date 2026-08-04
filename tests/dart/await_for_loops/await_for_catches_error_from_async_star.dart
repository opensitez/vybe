// vybe-test: dart/await_for_loops/await_for_catches_error_from_async_star
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

Stream<int> bad() async* {
  yield 1;
  throw Exception('fail');
}
Future<void> __vybeMain() async {
  var out = <String>[];
  try {
    await for (var v in bad()) { out.add('$v'); }
  } catch (_) {
    out.add('caught');
  }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,caught');
}
