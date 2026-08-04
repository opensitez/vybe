// vybe-test: dart/factory_constructors_deep/factory_cache_cleared_manually_still_works
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

class Session {
  static Session? _active;
  int id;
  Session._(this.id);
  factory Session(int id) {
    _active = Session._(id);
    return _active!;
  }
}
void __vybeMain() {
  __p(Session(3).id);
}

void main() {
  __vybeMain();
  __check('3');
}
