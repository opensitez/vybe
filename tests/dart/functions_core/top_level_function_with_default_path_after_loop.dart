// vybe-test: dart/functions_core/top_level_function_with_default_path_after_loop
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

int indexOf(List<int> xs, int target) {
  for (var i = 0; i < xs.length; i++) {
    if (xs[i] == target) {
      return i;
    }
  }
  return -1;
}
void __vybeMain() {
  __p(indexOf([10, 20, 30], 20));
  __p(indexOf([10, 20, 30], 99));
}

void main() {
  __vybeMain();
  __check('1\n-1');
}
