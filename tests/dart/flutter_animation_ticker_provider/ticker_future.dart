// vybe-test: dart/flutter_animation_ticker_provider/ticker_future
// origin: languages/dart/tests/dart/test_flutter_animation_ticker_provider.rs

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

import 'package:flutter/scheduler.dart';
void __vybeMain() async {
  final ticker = Ticker((_) {});
  final future = ticker.start();
  ticker.stop(canceled: true);
  try {
    await future;
    __p('finished');
  } catch(e) {
    __p('canceled'); // TickerCanceled
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('canceled');
}
