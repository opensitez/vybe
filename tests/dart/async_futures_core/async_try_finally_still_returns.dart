// vybe-test: dart/async_futures_core/async_try_finally_still_returns
// origin: languages/dart/tests/dart/test_async_futures_core.rs

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

Future<int> guarded() async {
  try {
    return 2;
  } finally {
    __p('cleanup');
  }
}
void __vybeMain() async {
  __p(await guarded());
}

Future<void> main() async {
  await __vybeMain();
  __check('cleanup\n2');
}
