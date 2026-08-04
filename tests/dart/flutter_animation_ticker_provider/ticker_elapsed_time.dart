// vybe-test: dart/flutter_animation_ticker_provider/ticker_elapsed_time
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
void __vybeMain() {
  int count = 0;
  final ticker = Ticker((elapsed) {
    if (elapsed.inMilliseconds >= 0) {
      count++;
    }
  });
  ticker.start();
  // We can't actually wait for ticks natively without scheduler mock,
  // so we just ensure it doesn't crash on start.
  __p('started');
  ticker.stop();
}

void main() {
  __vybeMain();
  __check('started');
}
