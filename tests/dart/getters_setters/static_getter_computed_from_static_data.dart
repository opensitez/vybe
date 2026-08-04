// vybe-test: dart/getters_setters/static_getter_computed_from_static_data
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class App {
  static List<String> _tags = ['dart'];
  static int get tagCount {
    return _tags.length;
  }
}
void __vybeMain() {
  __p(App.tagCount);
}

void main() {
  __vybeMain();
  __check('1');
}
