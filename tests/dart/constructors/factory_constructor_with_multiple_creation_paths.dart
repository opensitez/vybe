// vybe-test: dart/constructors/factory_constructor_with_multiple_creation_paths
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

class Num {
  int v;
  Num(this.v);
  factory Num.parse(String s) {
    if (s == 'zero') {
      return Num(0);
    }
    return Num(int.parse(s));
  }
}
void __vybeMain() {
  __p(Num.parse('zero').v);
}

void main() {
  __vybeMain();
  __check('0');
}
