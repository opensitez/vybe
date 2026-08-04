// vybe-test: dart/comparable_ordering/comparable_list_not_sorted_detected
// origin: languages/dart/tests/dart/test_comparable_ordering.rs

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

class Step implements Comparable<Step> {
  int s;
  Step(this.s);
  int compareTo(Step other) => s.compareTo(other.s);
}
bool isSorted(List<Step> list) {
  for (var i = 1; i < list.length; i++) {
    if (list[i - 1].compareTo(list[i]) > 0) return false;
  }
  return true;
}
void __vybeMain() {
  __p(isSorted([Step(3), Step(1)]));
}

void main() {
  __vybeMain();
  __check('false');
}
