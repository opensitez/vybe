// vybe-test: dart/dart_developer_metrics_gauges/metrics_counter_value
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
  final counter = Counter('my.counter2', 'description');
  counter.value = 10.5;
  __p(counter.value);
}

void main() {
  __vybeMain();
  __check('10.5');
}
