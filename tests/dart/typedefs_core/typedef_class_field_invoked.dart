// vybe-test: dart/typedefs_core/typedef_class_field_invoked
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef Handler = String Function(String);
class Greeter {
  Handler handler;
  Greeter(this.handler);
  String run(String name) {
    return handler(name);
  }
}
void __vybeMain() {
  var g = Greeter((name) => 'hi $name');
  __p(g.run('dart'));
}

void main() {
  __vybeMain();
  __check('hi dart');
}
