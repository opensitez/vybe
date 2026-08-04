// vybe-test: dart/operator_overloading/operator_equals_after_mutation_false
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

class Score {
  int pts;
  Score(this.pts);
  bool operator ==(Object other) {
    return other is Score && pts == other.pts;
  }
}
void __vybeMain() {
  var a = Score(1);
  var b = Score(1);
  b.pts = 2;
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('false');
}
