// vybe-test: dart/comparable_ordering/comparable_chained_compare_in_sort
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

class Item implements Comparable<Item> {
  int pri;
  int seq;
  Item(this.pri, this.seq);
  int compareTo(Item other) {
    var c = pri.compareTo(other.pri);
    return c != 0 ? c : seq.compareTo(other.seq);
  }
}
void __vybeMain() {
  var items = [Item(2, 1), Item(1, 2), Item(1, 1)];
  items.sort();
  __p(items[0].pri);
  __p(items[0].seq);
  __p(items[2].pri);
}

void main() {
  __vybeMain();
  __check('1\n1\n2');
}
