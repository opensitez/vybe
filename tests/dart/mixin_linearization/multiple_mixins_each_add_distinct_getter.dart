// vybe-test: dart/mixin_linearization/multiple_mixins_each_add_distinct_getter
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

mixin Ga {
  int get a {
    return 1;
  }
}
mixin Gb {
  int get b {
    return 2;
  }
}
mixin Gc {
  int get c {
    return 3;
  }
}
class All with Ga, Gb, Gc {}
void __vybeMain() {
  var x = All();
  __p(x.a + x.b + x.c);
}

void main() {
  __vybeMain();
  __check('6');
}
