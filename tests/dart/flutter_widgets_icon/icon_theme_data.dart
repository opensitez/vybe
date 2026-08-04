// vybe-test: dart/flutter_widgets_icon/icon_theme_data
// origin: languages/dart/tests/dart/test_flutter_widgets_icon.rs

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

import 'package:flutter/widgets.dart';
void __vybeMain() {
  final itd = const IconThemeData(color: Color(0xFF112233), size: 30.0);
  __p('${itd.color?.value == 0xFF112233}:${itd.size}');
}

void main() {
  __vybeMain();
  __check('true:30.0');
}
