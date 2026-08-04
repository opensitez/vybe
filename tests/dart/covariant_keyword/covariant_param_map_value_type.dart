// vybe-test: dart/covariant_keyword/covariant_param_map_value_type
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

class Registry {
  void register(Map<String, Object> m) {}
}
class IntRegistry extends Registry {
  @override
  void register(covariant Map<String, int> m) {
    __p(m['x']);
  }
}
void __vybeMain() {
  IntRegistry().register({'x': 9});
}

void main() {
  __vybeMain();
  __check('9');
}
