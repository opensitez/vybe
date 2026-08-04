// vybe-test: dart/dart_io_process_start_streams/process_start_detached_with_stdio
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('echo', ['detached'], mode: ProcessStartMode.detachedWithStdio);
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.trim());
}
