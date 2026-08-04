// vybe-test: dart/abstract_members/abstract_setter_triggers_side_effect
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

abstract class Store {
  set token(String t);
  String get token;
}
class MemStore extends Store {
  String _t = '';
  set token(String t) {
    _t = t;
  }
  String get token => _t;
}
void __vybeMain() {
  var s = MemStore();
  s.token = 'abc';
  __p(s.token.length);
}

void main() {
  __vybeMain();
  __check('3');
}
