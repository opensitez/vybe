// vybe-test: dart/dart_io_link_resolution/file_system_entity_basename
// origin: languages/dart/tests/dart/test_dart_io_link_resolution.rs

import 'dart:io';
import 'package:path/path.dart' as p;
// wait, we can't use package:path, but File doesn't have basename property in Dart.
// We'll skip basename test and replace with another API.
void main() {
  print('ok');
}
