// vybe-test: dart/functions_core/recursive_list_length_via_tail_style
// origin: languages/dart/tests/dart/test_functions_core.rs

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

int length(List<int> xs, int acc) {
  if (xs.isEmpty) {
    return acc;
  }
  return length(xs.sublist(1), acc + 1);
}
void __vybeMain() {
  __p(length([1, 2, 3, 4], 0));
}

void main() {
  __vybeMain();
  __check('4');
}
