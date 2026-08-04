// vybe-test: dart/future_combinators/future_catch_error_test_only_handles_matching
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
  try {
    await Future<int>.error('fail').catchError((e) => 1, test: (e) => e == 'other');
  } catch (e) {
    __p('err:$e');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('err:fail');
}
