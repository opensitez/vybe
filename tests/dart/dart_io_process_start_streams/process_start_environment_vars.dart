// vybe-test: dart/dart_io_process_start_streams/process_start_environment_vars
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('env', [], environment: {'CUSTOM_VAR': 'abc_123'});
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.contains('CUSTOM_VAR=abc_123'));
}
