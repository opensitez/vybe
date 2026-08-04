// vybe-test: dart/flutter_ui_path_metrics/path_transform
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
import 'dart:typed_data';
void __vybeMain() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  // Identity matrix is 16 elements
  final matrix = Float64List(16);
  matrix[0] = 2.0; // scale X
  matrix[5] = 2.0; // scale Y
  matrix[10] = 1.0;
  matrix[15] = 1.0;
  
  final transformed = path.transform(matrix);
  __p(transformed.getBounds().width);
}

void main() {
  __vybeMain();
  __check('20.0');
}
