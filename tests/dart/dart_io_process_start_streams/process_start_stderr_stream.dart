// vybe-test: dart/dart_io_process_start_streams/process_start_stderr_stream
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('ls', ['does_not_exist_xyz123']);
  final err = await process.stderr.transform(utf8.decoder).join();
  print(err.isNotEmpty);
}
