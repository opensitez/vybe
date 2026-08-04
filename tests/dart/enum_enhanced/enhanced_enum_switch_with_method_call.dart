// vybe-test: dart/enum_enhanced/enhanced_enum_switch_with_method_call
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Action {
  start,
  stop;
  String verb() {
    return name;
  }
}
String run(Action a) {
  switch (a) {
    case Action.start:
      return a.verb() + 'ed';
    case Action.stop:
      return a.verb() + 'ped';
  }
}
void __vybeMain() {
  __p(run(Action.start));
}

void main() {
  __vybeMain();
  __check('started');
}
