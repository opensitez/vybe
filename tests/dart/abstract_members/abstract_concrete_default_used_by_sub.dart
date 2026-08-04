// vybe-test: dart/abstract_members/abstract_concrete_default_used_by_sub
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Parser {
  String parse(String input) {
    return input.trim();
  }
  String format(String s);
}
class SimpleParser extends Parser {
  String format(String s) {
    return parse(s).toUpperCase();
  }
}
void __vybeMain() {
  __p(SimpleParser().format('  hi  '));
}

void main() {
  __vybeMain();
  __check('HI');
}
