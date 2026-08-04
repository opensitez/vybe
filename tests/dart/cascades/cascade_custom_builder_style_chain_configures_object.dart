// vybe-test: dart/cascades/cascade_custom_builder_style_chain_configures_object
// origin: languages/dart/tests/dart/test_cascades.rs

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

class Report {
  String title = '';
  int pages = 0;
  void setTitle(String t) { title = t; }
  void addPage() { pages += 1; }
}
void __vybeMain() {
  var doc = Report();
  doc..setTitle('Q1')..addPage()..addPage();
  __p(doc.title);
  __p(doc.pages);
}

void main() {
  __vybeMain();
  __check('Q1\n2');
}
