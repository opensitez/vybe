// vybe-test: dart/factory_constructors_deep/factory_from_json_with_fallback_name
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

class Label {
  String text;
  Label._(this.text);
  factory Label.fromJson(Map<String, dynamic> json) {
    var t = json['text'] ?? json['label'] ?? 'unknown';
    return Label._(t);
  }
}
void __vybeMain() {
  __p(Label.fromJson({'label': 'ok'}).text);
}

void main() {
  __vybeMain();
  __check('ok');
}
