// vybe-test: dart/factory_constructors_deep/factory_cached_by_string_key
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

class Icon {
  static final Map<String, Icon> _cache = {};
  String name;
  Icon._(this.name);
  factory Icon(String n) {
    return _cache.putIfAbsent(n, () => Icon._(n));
  }
}
void __vybeMain() {
  var a = Icon('home');
  var b = Icon('home');
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true');
}
