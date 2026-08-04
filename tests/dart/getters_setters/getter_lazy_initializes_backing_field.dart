// vybe-test: dart/getters_setters/getter_lazy_initializes_backing_field
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

class Lazy {
  String? _cache;
  String get label {
    _cache ??= 'ready';
    return _cache!;
  }
}
void __vybeMain() {
  __p(Lazy().label);
}

void main() {
  __vybeMain();
  __check('ready');
}
