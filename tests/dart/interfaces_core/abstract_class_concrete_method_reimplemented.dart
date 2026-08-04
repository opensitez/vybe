// vybe-test: dart/interfaces_core/abstract_class_concrete_method_reimplemented
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Base {
  String prefix() {
    return 'pre';
  }
  String suffix();
}
class Impl implements Base {
  String prefix() {
    return 'pre';
  }
  String suffix() {
    return 'suf';
  }
}
void __vybeMain() {
  var i = Impl();
  __p(i.prefix() + i.suffix());
}

void main() {
  __vybeMain();
  __check('presuf');
}
