// vybe-test: dart/super_parameters/super_param_base_two_named_sub_adds_positional
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Config {
  int port;
  String host;
  Config({this.port = 80, this.host = 'localhost'});
}
class AppConfig extends Config {
  String app;
  AppConfig(this.app, {super.port, super.host});
}
void __vybeMain() {
  var c = AppConfig('vybe', port: 8080);
  __p('${c.app}:${c.port}');
}

void main() {
  __vybeMain();
  __check('vybe:8080');
}
