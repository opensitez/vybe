// vybe-test: dart/flutter_ui_path_metrics/path_metrics_tangent_for_offset
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
  path.lineTo(10, 0);
  final metrics = path.computeMetrics().toList();
  final tangent = metrics[0].getTangentForOffset(5);
  __p('${tangent!.position.dx}:${tangent.vector.dx}');
}

void main() {
  __vybeMain();
  __check('5.0:1.0');
}
