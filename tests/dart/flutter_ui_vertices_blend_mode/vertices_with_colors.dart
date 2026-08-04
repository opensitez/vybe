// vybe-test: dart/flutter_ui_vertices_blend_mode/vertices_with_colors
// origin: languages/dart/tests/dart/test_flutter_ui_vertices_blend_mode.rs

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
  final v = Vertices(
    VertexMode.triangles,
    [Offset(0, 0), Offset(10, 0), Offset(0, 10)],
    colors: [Color(0xFFFF0000), Color(0xFF00FF00), Color(0xFF0000FF)]
  );
  __p(v != null);
}

void main() {
  __vybeMain();
  __check('true');
}
