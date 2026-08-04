// vybe-test: dart/factory_constructors_deep/factory_from_json_nested_map
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

class Address {
  String city;
  Address._(this.city);
  factory Address.fromJson(Map<String, dynamic> json) {
    return Address._(json['city']);
  }
}
class Person {
  String name;
  Address addr;
  Person._(this.name, this.addr);
  factory Person.fromJson(Map<String, dynamic> json) {
    return Person._(json['name'], Address.fromJson(json['addr']));
  }
}
void __vybeMain() {
  var p = Person.fromJson({'name': 'Eve', 'addr': {'city': 'Paris'}});
  __p(p.addr.city);
}

void main() {
  __vybeMain();
  __check('Paris');
}
