// vybe-test: dart/classes_core/override_setter_updates_subclass_storage
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Base {
  int _v = 0;
  set val(int n) {
    _v = n;
  }
  int get val {
    return _v;
  }
}
class Child extends Base {
  @override
  set val(int n) {
    _v = n * 2;
  }
}
void __vybeMain() {
  var c = Child();
  c.val = 3;
  __p(c.val);
}

void main() {
  __vybeMain();
  __check('6');
}
