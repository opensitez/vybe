// vybe-test: dart/future_combinators/future_then_catch_error_when_complete_execution_order
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
  await Future.value(1)
      .then((v) {
        log.add('then');
        return v;
      })
      .catchError((e) {
        log.add('catch');
        return 0;
      })
      .whenComplete(() => log.add('complete'));
  __p(log.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('then,complete');
}
