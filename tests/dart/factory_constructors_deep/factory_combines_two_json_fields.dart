// vybe-test: dart/factory_constructors_deep/factory_combines_two_json_fields
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class FullName {
  String full;
  FullName._(this.full);
  factory FullName.fromJson(Map<String, dynamic> json) {
    return FullName._('${json['first']} ${json['last']}');
  }
}
void __vybeMain() {
  __p(FullName.fromJson({'first': 'Ada', 'last': 'Lovelace'}).full);
}

void main() {
  __vybeMain();
  __check('Ada Lovelace');
}
