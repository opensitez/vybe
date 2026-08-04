// vybe-test: dart/dart_io_process_start_streams/process_start_working_directory
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
import 'dart:convert';
void main() async {
  final dir = Directory.systemTemp.createTempSync('start_wd_');
  File('${dir.path}/wd_file.txt').createSync();
  final process = await Process.start('ls', [], workingDirectory: dir.path);
  final out = await process.stdout.transform(utf8.decoder).join();
  print(out.contains('wd_file.txt'));
  dir.deleteSync(recursive: true);
}
