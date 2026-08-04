// vybe-test: dart/mixins_core/mixin_with_static_host_class_field
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

class Host {
  static int global = 7;
}
mixin ReadGlobal on Host {
  int read() {
    return global;
  }
}
class App extends Host with ReadGlobal {}
void __vybeMain() {
  __p(App().read());
}

void main() {
  __vybeMain();
  __check('7');
}
