// vybe-test: dart/dart_io_directory_symlinks/link_resolve_symbolic_links_sync_on_link
// origin: languages/dart/tests/dart/test_dart_io_directory_symlinks.rs

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
  final file = File('resolve_target.txt');
  file.createSync();
  final link = Link('resolve_link.lnk');
  link.createSync(file.path);
  final resolved = link.resolveSymbolicLinksSync();
  __p(resolved.isNotEmpty);
  link.deleteSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
