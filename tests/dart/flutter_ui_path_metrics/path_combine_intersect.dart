// vybe-test: dart/flutter_ui_path_metrics/path_combine_intersect
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
  final p1 = Path()..addRect(Rect.fromLTRB(0, 0, 10, 10));
  final p2 = Path()..addRect(Rect.fromLTRB(5, 5, 15, 15));
  final combined = Path.combine(PathOperation.intersect, p1, p2);
  __p('${combined.getBounds().width}:${combined.getBounds().height}');
}

void main() {
  __vybeMain();
  __check('5.0:5.0');
}
