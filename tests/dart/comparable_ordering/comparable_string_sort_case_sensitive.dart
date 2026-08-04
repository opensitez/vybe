// vybe-test: dart/comparable_ordering/comparable_string_sort_case_sensitive
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

class Name implements Comparable<Name> {
  String s;
  Name(this.s);
  int compareTo(Name other) => s.compareTo(other.s);
}
void __vybeMain() {
  __p(Name('B').compareTo(Name('a')));
}

void main() {
  __vybeMain();
  __check('-1');
}
