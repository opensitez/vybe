// vybe-test: dart/regexp_core/regexp_first_match_domain_capture
// origin: languages/dart/tests/dart/test_regexp_core.rs

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

void __vybeMain() {
  var re = RegExp(r'(\w+)@(\w+)');
  var m = re.firstMatch('user@host');
  __p(m!.group(2));
}

void main() {
  __vybeMain();
  __check('host');
}
