// vybe-test: dart/dart_io_process_start_streams/process_start_stdin_add
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('cat', []);
  process.stdin.add(utf8.encode('byte_data\n'));
  await process.stdin.close();
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.trim());
}
