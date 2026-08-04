// vybe-test: dart/flutter_material_bottom_sheet/bottom_sheet_animation_controller
// origin: languages/dart/tests/dart/test_flutter_material_bottom_sheet.rs

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

import 'package:flutter/material.dart';
void __vybeMain() {
  final bs = BottomSheet(
    animationController: AnimationController(
      vsync: const TestVSync(),
      duration: const Duration(seconds: 1),
    ),
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  __p(bs.animationController != null);
}

class TestVSync implements TickerProvider {
  const TestVSync();
  @override
  Ticker createTicker(TickerCallback onTick) => Ticker(onTick);
}

void main() {
  __vybeMain();
  __check('true');
}
