// vybe-test: dart/class_modifiers/sealed_switch_with_object_pattern_field
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Response {}
class Success extends Response {
  int code;
  Success(this.code);
}
class Failure extends Response {
  String reason;
  Failure(this.reason);
}
String describe(Response r) {
  switch (r) {
    case Success(code: 200):
      return 'ok';
    case Success(code: var c):
      return 'code:$c';
    case Failure(reason: var msg):
      return msg;
  }
}
void __vybeMain() {
  __p(describe(Success(200)));
}

void main() {
  __vybeMain();
  __check('ok');
}
