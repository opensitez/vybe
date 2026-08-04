// vybe-test: dart/named_parameters/constructor_named_reordered
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

class Pair {
  final int first;
  final int second;
  Pair({required this.first, required this.second});
}
void __vybeMain() {
  var p = Pair(second: 2, first: 1);
  __p('${p.first},${p.second}');
}

void main() {
  __vybeMain();
  __check('1,2');
}
