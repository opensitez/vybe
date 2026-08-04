// vybe-test: dart/flutter_widgets_flexible/flexible_apply_parent_data_flex
// origin: languages/dart/tests/dart/test_flutter_widgets_flexible.rs

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
import 'package:flutter/rendering.dart';
void __vybeMain() {
  final f = Flexible(flex: 3, fit: FlexFit.tight, child: const SizedBox());
  __p('compiles');
}

void main() {
  __vybeMain();
  __check('compiles');
}
