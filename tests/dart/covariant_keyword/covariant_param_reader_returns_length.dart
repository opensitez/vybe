// vybe-test: dart/covariant_keyword/covariant_param_reader_returns_length
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

class Reader {
  int size(List<Object> buf) {
    return buf.length;
  }
}
class ByteReader extends Reader {
  @override
  int size(covariant List<int> buf) {
    return buf.length * 2;
  }
}
void __vybeMain() {
  __p(ByteReader().size([1, 2, 3, 4]));
}

void main() {
  __vybeMain();
  __check('8');
}
