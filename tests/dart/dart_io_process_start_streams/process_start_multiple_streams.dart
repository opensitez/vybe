// vybe-test: dart/dart_io_process_start_streams/process_start_multiple_streams
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

import 'dart:io';
void main() async {
  final process = await Process.start('echo', ['test']);
  int count = 0;
  process.stdout.listen((_) { count++; });
  process.stderr.listen((_) {});
  await process.exitCode;
  print(count > 0);
}
