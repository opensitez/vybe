// vybe-test: dart/comparable_ordering/comparable_sort_ascending_list
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

class Score implements Comparable<Score> {
  int pts;
  Score(this.pts);
  int compareTo(Score other) => pts.compareTo(other.pts);
}
void __vybeMain() {
  var list = [Score(30), Score(10), Score(20)];
  list.sort();
  __p(list[0].pts);
  __p(list[2].pts);
}

void main() {
  __vybeMain();
  __check('10\n30');
}
