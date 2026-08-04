// vybe-test: dart/dart_developer_metrics_gauges/metrics_gauge_creation
// origin: languages/dart/tests/dart/test_dart_developer_metrics_gauges.rs

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

import 'dart:developer';
void __vybeMain() {
  final gauge = Gauge('my.gauge', 'A test gauge', min: 0.0, max: 100.0);
  __p(gauge.name);
  __p(gauge.min);
  __p(gauge.max);
}

void main() {
  __vybeMain();
  __check('my.gauge\n0.0\n100.0');
}
