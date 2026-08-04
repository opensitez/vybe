// vybe-test: dart/interfaces_core/abstract_interface_with_multiple_methods
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

abstract class Repo {
  void save(String k);
  String load(String k);
}
class MapRepo implements Repo {
  String store = '';
  void save(String k) {
    store = k;
  }
  String load(String k) {
    return store;
  }
}
void __vybeMain() {
  var r = MapRepo();
  r.save('key');
  __p(r.load('key'));
}

void main() {
  __vybeMain();
  __check('key');
}
