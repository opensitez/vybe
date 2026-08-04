// vybe-test: dart/functions_core/nested_local_functions_three_levels
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

void __vybeMain() {
  int outer(int x) {
    int middle(int y) {
      int inner(int z) {
        return x + y + z;
      }
      return inner(3);
    }
    return middle(2);
  }
  __p(outer(1));
}

void main() {
  __vybeMain();
  __check('6');
}
