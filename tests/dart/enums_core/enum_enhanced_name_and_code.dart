// vybe-test: dart/enums_core/enum_enhanced_name_and_code
// origin: languages/dart/tests/dart/test_enums_core.rs

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
  medium(5),
  high(10);
  final int weight;
  const Priority(this.weight);
}
void __vybeMain() {
  var p = Priority.medium;
  __p('${p.name}:${p.weight}');
}

void main() {
  __vybeMain();
  __check('medium:5');
}
