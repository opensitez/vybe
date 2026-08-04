// vybe-test: dart/enum_enhanced/enhanced_enum_field_compare_in_method
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Priority {
  low(1),
  high(5);
  final int weight;
  const Priority(this.weight);
  bool beats(Priority other) {
    return weight > other.weight;
  }
}
void __vybeMain() {
  __p(Priority.high.beats(Priority.low));
}

void main() {
  __vybeMain();
  __check('true');
}
