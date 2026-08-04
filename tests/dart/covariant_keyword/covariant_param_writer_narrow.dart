// vybe-test: dart/covariant_keyword/covariant_param_writer_narrow
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Writer {
  void write(Object data) {}
}
class TextWriter extends Writer {
  @override
  void write(covariant String data) {
    __p(data.length);
  }
}
void __vybeMain() {
  TextWriter().write('abc');
}

void main() {
  __vybeMain();
  __check('3');
}
