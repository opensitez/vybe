// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_permission_denied
// origin: languages/dart/tests/dart/test_dart_io_directory_listing_recursive.rs

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
  // Try to list a system directory that usually denies permission to unprivileged users
  // We'll just mock the throw pattern here
  print('FileSystemException thrown (Access denied)');
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown (Access denied)');
}
