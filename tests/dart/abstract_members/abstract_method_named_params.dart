// vybe-test: dart/abstract_members/abstract_method_named_params
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

abstract class Connect {
  bool open({required String host, int port = 80});
}
class Tcp extends Connect {
  bool open({required String host, int port = 80}) {
    __p('$host:$port');
    return true;
  }
}
void __vybeMain() {
  __p(Tcp().open(host: 'localhost', port: 8080));
}

void main() {
  __vybeMain();
  __check('localhost:8080\ntrue');
}
