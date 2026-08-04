// vybe-test: dart/dart_io_link_resolution/resolve_symbolic_links_sync_nested_links
// origin: languages/dart/tests/dart/test_dart_io_link_resolution.rs

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

import 'dart:io';
void __vybeMain() {
  final file = File('nested_tgt.txt');
  file.createSync();
  final l1 = Link('l1.lnk');
  final l2 = Link('l2.lnk');
  l1.createSync(file.path);
  l2.createSync(l1.path); // l2 -> l1 -> file
  
  final resolved = l2.resolveSymbolicLinksSync();
  __p(resolved == file.absolute.path);
  
  l2.deleteSync();
  l1.deleteSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
