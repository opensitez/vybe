// vybe-test: dart/callable_objects/call_returns_list
// origin: languages/dart/tests/dart/test_callable_objects.rs

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

class Range {
  List<int> call(int start, int end) {
    var out = <int>[];
    for (var i = start; i <= end; i++) {
      out.add(i);
    }
    return out;
  }
}
void __vybeMain() {
  __p(Range()(1, 3).join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
