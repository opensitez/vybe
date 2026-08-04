// vybe-test: dart/named_parameters/constructor_required_named_params
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

class Server {
  final String host;
  final int port;
  Server({required this.host, required this.port});
}
void __vybeMain() {
  var s = Server(host: '127.0.0.1', port: 9000);
  __p('${s.host}:${s.port}');
}

void main() {
  __vybeMain();
  __check('127.0.0.1:9000');
}
