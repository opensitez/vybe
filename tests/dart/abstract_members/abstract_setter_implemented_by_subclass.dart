// vybe-test: dart/abstract_members/abstract_setter_implemented_by_subclass
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Mutable {
  set value(int v);
  int get value;
}
class Counter extends Mutable {
  int _v = 0;
  set value(int v) {
    _v = v;
  }
  int get value => _v;
}
void __vybeMain() {
  var c = Counter();
  c.value = 9;
  __p(c.value);
}

void main() {
  __vybeMain();
  __check('9');
}
