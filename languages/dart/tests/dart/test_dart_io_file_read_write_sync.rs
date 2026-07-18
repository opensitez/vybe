use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io File Read/Write Sync
// ═══════════════════════════════════════════════════════════

#[test]
fn file_read_as_string_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_read_sync.txt');
  file.writeAsStringSync('Hello Dart IO');
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["Hello Dart IO"]
    );
}

#[test]
fn file_read_as_bytes_sync_large() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_bytes_sync.bin');
  file.writeAsBytesSync([0, 255, 128, 64]);
  final bytes = file.readAsBytesSync();
  print('${bytes.length}:${bytes[1]}');
  file.deleteSync();
}
"#
        ),
        vec!["4:255"]
    );
}

#[test]
fn file_write_as_string_sync_overwrite() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_overwrite_sync.txt');
  file.writeAsStringSync('Initial Data');
  file.writeAsStringSync('New Data', mode: FileMode.write);
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["New Data"]
    );
}

#[test]
fn file_write_as_string_sync_append() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_append_sync.txt');
  file.writeAsStringSync('Part 1, ');
  file.writeAsStringSync('Part 2', mode: FileMode.append);
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["Part 1, Part 2"]
    );
}

#[test]
fn file_write_as_bytes_sync_append() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_append_bytes_sync.bin');
  file.writeAsBytesSync([10, 20]);
  file.writeAsBytesSync([30, 40], mode: FileMode.append);
  final bytes = file.readAsBytesSync();
  print(bytes.join('-'));
  file.deleteSync();
}
"#
        ),
        vec!["10-20-30-40"]
    );
}

#[test]
fn file_read_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('does_not_exist_123.txt');
  try {
    file.readAsStringSync();
    print('Failed to throw');
  } on FileSystemException catch (e) {
    print('FileSystemException thrown');
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn file_write_read_only_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print('FileSystemException thrown on read-only');
}
"#
        ),
        vec!["FileSystemException thrown on read-only"]
    );
}

#[test]
fn file_read_as_lines_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_lines_sync.txt');
  file.writeAsStringSync('Line1\nLine2\r\nLine3');
  final lines = file.readAsLinesSync();
  print('${lines.length}:${lines[0]}:${lines[2]}');
  file.deleteSync();
}
"#
        ),
        vec!["3:Line1:Line3"]
    );
}

#[test]
fn file_write_empty_string() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_empty_sync.txt');
  file.writeAsStringSync('');
  print(file.lengthSync());
  file.deleteSync();
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn file_read_explicit_utf8_encoding() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
import 'dart:convert';
void main() {
  final file = File('test_utf8_sync.txt');
  file.writeAsBytesSync(utf8.encode('Därt'));
  print(file.readAsStringSync(encoding: utf8));
  file.deleteSync();
}
"#
        ),
        vec!["Därt"]
    );
}

#[test]
fn file_read_explicit_latin1_encoding() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
import 'dart:convert';
void main() {
  final file = File('test_latin1_sync.txt');
  file.writeAsBytesSync(latin1.encode('Därt'));
  print(file.readAsStringSync(encoding: latin1));
  file.deleteSync();
}
"#
        ),
        vec!["Därt"]
    );
}

#[test]
fn file_write_with_flush() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_flush_sync.txt');
  file.writeAsStringSync('Flushed data', flush: true);
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["Flushed data"]
    );
}

#[test]
fn file_write_bytes_with_flush() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_flush_bytes_sync.bin');
  file.writeAsBytesSync([100, 200], flush: true);
  print(file.readAsBytesSync().length);
  file.deleteSync();
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn file_write_to_root_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('/');
  try {
    file.writeAsStringSync('test');
  } on FileSystemException {
    print('FileSystemException thrown');
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn file_concurrent_read_lock() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_lock.txt');
  file.writeAsStringSync('locked');
  final raf = file.openSync(mode: FileMode.read);
  raf.lockSync(FileLock.exclusive);
  print('Locked successfully');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["Locked successfully"]
    );
}

#[test]
fn file_concurrent_write_lock() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_write_lock.txt');
  file.writeAsStringSync('write lock');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  print('Write locked successfully');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["Write locked successfully"]
    );
}

#[test]
fn file_open_sync_write_only() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_open_write.txt');
  final raf = file.openSync(mode: FileMode.writeOnly);
  raf.writeStringSync('Only writing');
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["Only writing"]
    );
}

#[test]
fn file_open_sync_append() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_open_append.txt');
  file.writeAsStringSync('Start-');
  final raf = file.openSync(mode: FileMode.append);
  raf.writeStringSync('End');
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["Start-End"]
    );
}

#[test]
fn directory_as_file_read_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp;
  final file = File(dir.path);
  try {
    file.readAsStringSync();
  } on FileSystemException {
    print('FileSystemException thrown on dir read');
  }
}
"#
        ),
        vec!["FileSystemException thrown on dir read"]
    );
}

#[test]
fn directory_as_file_write_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp;
  final file = File(dir.path);
  try {
    file.writeAsStringSync('test');
  } on FileSystemException {
    print('FileSystemException thrown on dir write');
  }
}
"#
        ),
        vec!["FileSystemException thrown on dir write"]
    );
}
