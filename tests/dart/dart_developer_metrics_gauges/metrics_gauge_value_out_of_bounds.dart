// vybe-test: dart/dart_developer_metrics_gauges/metrics_gauge_value_out_of_bounds
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
  final gauge = Gauge('my.gauge3', 'desc', min: 0.0, max: 10.0);
  try {
    // Some implementations might clamp, throw, or accept it. 
    // Usually Dart just accepts it but tools might warn.
    gauge.value = 20.0;
    __p(gauge.value);
  } catch(e) {
    __p('ArgumentError'); // In case it strictly throws
  }
}

void main() {
  __vybeMain();
  __check('20.0');
}
