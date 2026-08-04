// vybe-test: dart/covariant_keyword/covariant_param_error_type_narrow
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

class ErrorBase {}
class NetworkError extends ErrorBase {
  int code;
  NetworkError(this.code);
}
class Handler {
  void handle(ErrorBase e) {}
}
class NetHandler extends Handler {
  @override
  void handle(covariant NetworkError e) {
    __p(e.code);
  }
}
void __vybeMain() {
  NetHandler().handle(NetworkError(404));
}

void main() {
  __vybeMain();
  __check('404');
}
