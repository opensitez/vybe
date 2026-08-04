// vybe-test: dart/enum_enhanced/enhanced_enum_switch_assign_var
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

enum Flag { on, off }
void __vybeMain() {
  var f = Flag.on;
  var result = '';
  switch (f) {
    case Flag.on:
      result = 'yes';
      break;
    case Flag.off:
      result = 'no';
      break;
  }
  __p(result);
}

void main() {
  __vybeMain();
  __check('yes');
}
