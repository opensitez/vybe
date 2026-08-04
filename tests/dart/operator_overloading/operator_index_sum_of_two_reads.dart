// vybe-test: dart/operator_overloading/operator_index_sum_of_two_reads
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Duo {
  List<int> pair;
  Duo(this.pair);
  int operator [](int i) {
    return pair[i];
  }
}
void __vybeMain() {
  var d = Duo([3, 7]);
  __p(d[0] + d[1]);
}

void main() {
  __vybeMain();
  __check('10');
}
