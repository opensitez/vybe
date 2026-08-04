// vybe-test: dart/abstract_members/abstract_implement_two_abstract_methods
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

abstract class RW {
  String read();
  void write(String s);
}
class Buffer extends RW {
  String data = '';
  String read() {
    return data;
  }
  void write(String s) {
    data = s;
  }
}
void __vybeMain() {
  var b = Buffer();
  b.write('data');
  __p(b.read());
}

void main() {
  __vybeMain();
  __check('data');
}
