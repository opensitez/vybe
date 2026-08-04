// vybe-test: dart/dart_io_process_start_streams/process_start_stdout_stream
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('echo', ['hello_stream']);
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.trim());
}
