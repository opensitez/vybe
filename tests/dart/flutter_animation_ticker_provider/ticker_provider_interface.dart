// vybe-test: dart/flutter_animation_ticker_provider/ticker_provider_interface
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
class MyProvider implements TickerProvider {
  @override
  Ticker createTicker(TickerCallback onTick) => Ticker(onTick);
}
void __vybeMain() {
  final p = MyProvider();
  final t = p.createTicker((_) {});
  __p(t != null);
}

void main() {
  __vybeMain();
  __check('true');
}
