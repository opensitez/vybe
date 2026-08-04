// vybe-test: dart/flutter_ui_paint_shader/paint_image_filter_matrix
// origin: languages/dart/tests/dart/test_flutter_ui_paint_shader.rs

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
  final paint = Paint();
  final matrix = Float64List(16);
  matrix[0] = 1.0; matrix[5] = 1.0; matrix[10] = 1.0; matrix[15] = 1.0;
  paint.imageFilter = ImageFilter.matrix(matrix);
  __p(paint.imageFilter != null);
}

void main() {
  __vybeMain();
  __check('true');
}
