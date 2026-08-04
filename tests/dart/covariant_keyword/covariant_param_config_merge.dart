// vybe-test: dart/covariant_keyword/covariant_param_config_merge
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

class Config {
  void merge(Map<String, Object> opts) {}
}
class AppConfig extends Config {
  int port = 80;
  @override
  void merge(covariant Map<String, int> opts) {
    if (opts.containsKey('port')) {
      port = opts['port']!;
    }
  }
}
void __vybeMain() {
  var c = AppConfig();
  c.merge({'port': 3000});
  __p(c.port);
}

void main() {
  __vybeMain();
  __check('3000');
}
