// vybe-test: dart/dart_io_process_start_streams/process_start_stdin_write
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  // Use cat to echo stdin to stdout
  final process = await Process.start('cat', []);
  process.stdin.writeln('input_data');
  await process.stdin.flush();
  await process.stdin.close(); // Need to close stdin so cat terminates
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.trim());
}
