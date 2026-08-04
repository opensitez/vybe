// vybe-test: dart/comparable_ordering/comparable_binary_search_style
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

class Slot implements Comparable<Slot> {
  int s;
  Slot(this.s);
  int compareTo(Slot other) => s.compareTo(other.s);
}
int indexOf(List<Slot> list, Slot target) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].compareTo(target) == 0) return i;
  }
  return -1;
}
void __vybeMain() {
  var list = [Slot(10), Slot(20), Slot(30)];
  __p(indexOf(list, Slot(20)));
}

void main() {
  __vybeMain();
  __check('1');
}
