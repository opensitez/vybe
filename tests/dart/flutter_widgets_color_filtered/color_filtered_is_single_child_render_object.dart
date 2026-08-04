// vybe-test: dart/flutter_widgets_color_filtered/color_filtered_is_single_child_render_object
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
    colorFilter: const ColorFilter.mode(Color(0xFF0000FF), BlendMode.srcOver),
  );
  __p(cf is SingleChildRenderObjectWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
