// vybe-test: dart/operator_overloading/operator_index_assign_then_read
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Buffer {
  List<int> buf;
  Buffer(this.buf);
  int operator [](int i) => buf[i];
  void operator []=(int i, int v) {
    buf[i] = v;
  }
}
void __vybeMain() {
  var b = Buffer([0, 0]);
  b[0] = 42;
  __p(b[0]);
}

void main() {
  __vybeMain();
  __check('42');
}
