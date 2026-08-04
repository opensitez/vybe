// vybe-test: dart/constructors/factory_with_static_state
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Seq {
  static int _n = 0;
  int id;
  Seq(this.id);
  factory Seq.next() {
    _n = _n + 1;
    return Seq(_n);
  }
}
void __vybeMain() {
  __p(Seq.next().id);
}

void main() {
  __vybeMain();
  __check('1');
}
