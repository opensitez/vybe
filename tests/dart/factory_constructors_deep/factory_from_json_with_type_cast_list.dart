// vybe-test: dart/factory_constructors_deep/factory_from_json_with_type_cast_list
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

class Pack {
  List<String> tags;
  Pack._(this.tags);
  factory Pack.fromJson(Map<String, dynamic> json) {
    var raw = json['tags'] as List;
    return Pack._(raw.map((e) => e as String).toList());
  }
}
void __vybeMain() {
  var p = Pack.fromJson({'tags': ['a', 'b']});
  __p(p.tags.join('-'));
}

void main() {
  __vybeMain();
  __check('a-b');
}
