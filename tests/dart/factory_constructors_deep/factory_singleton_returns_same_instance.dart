// vybe-test: dart/factory_constructors_deep/factory_singleton_returns_same_instance
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Logger {
  static Logger? _inst;
  int hits = 0;
  Logger._();
  factory Logger() {
    _inst ??= Logger._();
    return _inst!;
  }
}
void __vybeMain() {
  var a = Logger();
  var b = Logger();
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true');
}
