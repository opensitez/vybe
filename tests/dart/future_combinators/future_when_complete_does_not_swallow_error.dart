// vybe-test: dart/future_combinators/future_when_complete_does_not_swallow_error
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

Future<void> __vybeMain() async {
  var log = <String>[];
  try {
    await Future<int>.error('fail').whenComplete(() => log.add('wc'));
  } catch (e) {
    log.add('thrown');
  }
  __p(log.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('wc,thrown');
}
