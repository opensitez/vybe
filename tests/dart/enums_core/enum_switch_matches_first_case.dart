// vybe-test: dart/enums_core/enum_switch_matches_first_case
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Tier { bronze, silver, gold }
void __vybeMain() {
  var t = Tier.bronze;
  switch (t) {
    case Tier.bronze:
      __p('third');
      break;
    case Tier.silver:
      __p('second');
      break;
    default:
      __p('first');
  }
}

void main() {
  __vybeMain();
  __check('third');
}
