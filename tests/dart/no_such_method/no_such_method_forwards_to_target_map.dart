// vybe-test: dart/no_such_method/no_such_method_forwards_to_target_map
// origin: languages/dart/tests/dart/test_no_such_method.rs

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

class Forwarder {
  final Map<String, dynamic> target;
  Forwarder(this.target);
  @override
  dynamic noSuchMethod(Invocation inv) {
    var name = inv.memberName.toString();
    if (name.contains('get')) {
      var key = name.replaceAll('Symbol(\"', '').replaceAll('\")', '');
      return target[key];
    }
    return null;
  }
}
void __vybeMain() {
  dynamic f = Forwarder({'x': 7});
  __p(f.x);
}

void main() {
  __vybeMain();
  __check('null');
}
