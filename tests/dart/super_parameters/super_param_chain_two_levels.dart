// vybe-test: dart/super_parameters/super_param_chain_two_levels
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Root {
  int v;
  Root(this.v);
}
class Mid extends Root {
  Mid(super.v);
}
class Leaf extends Mid {
  Leaf(super.v);
}
void __vybeMain() {
  __p(Leaf(11).v);
}

void main() {
  __vybeMain();
  __check('11');
}
