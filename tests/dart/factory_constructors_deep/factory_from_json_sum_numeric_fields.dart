// vybe-test: dart/factory_constructors_deep/factory_from_json_sum_numeric_fields
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Totals {
  int sum;
  Totals._(this.sum);
  factory Totals.fromJson(Map<String, dynamic> json) {
    return Totals._(json['a'] + json['b']);
  }
}
void __vybeMain() {
  __p(Totals.fromJson({'a': 10, 'b': 5}).sum);
}

void main() {
  __vybeMain();
  __check('15');
}
