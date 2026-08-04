// vybe-test: dart/extensions/list_method_contains_all_values
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension ListHas on List<int> {
  bool containsAllExt(List<int> others) {
    for (var v in others) {
      if (!contains(v)) return false;
    }
    return true;
  }
}
void __vybeMain() {
  __p([1, 2, 3, 4].containsAllExt([2, 4]));
}

void main() {
  __vybeMain();
  __check('true');
}
