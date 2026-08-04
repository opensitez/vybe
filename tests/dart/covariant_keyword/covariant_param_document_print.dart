// vybe-test: dart/covariant_keyword/covariant_param_document_print
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

class Document {}
class Pdf extends Document {
  int pages;
  Pdf(this.pages);
}
class Printer {
  void printDoc(Document d) {}
}
class PdfPrinter extends Printer {
  @override
  void printDoc(covariant Pdf d) {
    __p(d.pages);
  }
}
void __vybeMain() {
  PdfPrinter().printDoc(Pdf(10));
}

void main() {
  __vybeMain();
  __check('10');
}
