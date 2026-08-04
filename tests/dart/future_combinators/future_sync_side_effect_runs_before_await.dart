// vybe-test: dart/future_combinators/future_sync_side_effect_runs_before_await
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
  var v = await Future.sync(() {
    log.add('work');
    return 1;
  });
  log.add('after');
  __p('${v}|${log.join(',')}');
}

Future<void> main() async {
  await __vybeMain();
  __check('1|work,after');
}
