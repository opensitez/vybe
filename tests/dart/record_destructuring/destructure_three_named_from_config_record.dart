// vybe-test: dart/record_destructuring/destructure_three_named_from_config_record
// origin: languages/dart/tests/dart/test_record_destructuring.rs

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

void __vybeMain() {
  var (host: h, port: p, tls: t) = (host: 'localhost', port: 8080, tls: false);
  __p(h);
  __p(p);
  __p(t);
}

void main() {
  __vybeMain();
  __check('localhost\n8080\nfalse');
}
