// vybe-test: dart/dart_io_directory_symlinks/link_stat_sync_without_following
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
  final link = Link('stat_link.lnk');
  link.createSync('non_existent_target.txt');
  // FileStat.statSync(path) follows links by default unless told not to,
  // but link.statSync doesn't follow link? Wait, Link doesn't have statSync natively in Dart.
  // It has FileStat.statSync(path).
  final stat = FileStat.statSync(link.path);
  // If followed, it's notFound. If not followed (not default), it's link.
  __p(stat.type == FileSystemEntityType.notFound);
  link.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
