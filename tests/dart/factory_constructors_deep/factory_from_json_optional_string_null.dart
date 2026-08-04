// vybe-test: dart/factory_constructors_deep/factory_from_json_optional_string_null
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

class Note {
  String? body;
  Note._(this.body);
  factory Note.fromJson(Map<String, dynamic> json) {
    return Note._(json['body']);
  }
}
void __vybeMain() {
  __p(Note.fromJson({}).body);
}

void main() {
  __vybeMain();
  __check('null');
}
