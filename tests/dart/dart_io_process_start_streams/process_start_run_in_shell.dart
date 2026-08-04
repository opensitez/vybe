// vybe-test: dart/dart_io_process_start_streams/process_start_run_in_shell
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final process = await Process.start('echo', ['in_shell'], runInShell: true);
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.trim());
}
