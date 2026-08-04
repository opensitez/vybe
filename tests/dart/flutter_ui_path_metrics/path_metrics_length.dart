// vybe-test: dart/flutter_ui_path_metrics/path_metrics_length
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
  __p(metrics.length == 1);
  __p(metrics[0].length);
}

void main() {
  __vybeMain();
  __check('true\n10.0');
}
