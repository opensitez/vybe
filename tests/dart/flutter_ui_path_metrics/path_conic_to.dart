// vybe-test: dart/flutter_ui_path_metrics/path_conic_to
// origin: languages/dart/tests/dart/test_flutter_ui_path_metrics.rs

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

import 'dart:ui';
void __vybeMain() {
  final path = Path();
  path.moveTo(0, 0);
  path.conicTo(5, 10, 10, 0, 1.0);
  __p(path.getBounds().height > 0);
}

void main() {
  __vybeMain();
  __check('true');
}
