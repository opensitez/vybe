// vybe-test: dart/comparable_ordering/comparable_equal_via_compare_to_zero
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

class Tag implements Comparable<Tag> {
  String label;
  Tag(this.label);
  int compareTo(Tag other) => label.compareTo(other.label);
}
void __vybeMain() {
  __p(Tag('a').compareTo(Tag('a')) == 0);
}

void main() {
  __vybeMain();
  __check('true');
}
