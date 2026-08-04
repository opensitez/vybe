// vybe-test: dart/flutter_widgets_color_filtered/color_filtered_creation
// origin: languages/dart/tests/dart/test_flutter_widgets_color_filtered.rs

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
  final cf = ColorFiltered(
    colorFilter: const ColorFilter.mode(Color(0xFFFF0000), BlendMode.srcIn),
    child: const SizedBox(),
  );
  __p(cf.colorFilter != null);
}

void main() {
  __vybeMain();
  __check('true');
}
