// vybe-test: dart/dart_io_process_run/process_run_sync_stderr_capture
// origin: languages/dart/tests/dart/test_dart_io_process_run.rs

import 'dart:io';
void main() {
  // Command that writes to stderr. 'ls' on non-existent file.
  final result = Process.runSync('ls', ['non_existent_file_xyz_123']);
  print(result.exitCode != 0);
  print((result.stderr as String).isNotEmpty);
}
