// vybe-test: dart/factory_constructors_deep/factory_from_json_with_default_for_missing_key
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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
  int timeout;
  Config._(this.timeout);
  factory Config.fromJson(Map<String, dynamic> json) {
    var t = json['timeout'];
    return Config._(t ?? 30);
  }
}
void __vybeMain() {
  __p(Config.fromJson({}).timeout);
}

void main() {
  __vybeMain();
  __check('30');
}
