use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io RandomAccessFile
// ═══════════════════════════════════════════════════════════

#[test]
fn random_access_file_write_string_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_write_string.txt');
  final raf = file.openSync(mode: FileMode.write);
  raf.writeStringSync('hello');
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn random_access_file_write_byte_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_write_byte.bin');
  final raf = file.openSync(mode: FileMode.write);
  raf.writeByteSync(65); // 'A'
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["A"]
    );
}

#[test]
fn random_access_file_write_from_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_write_from.bin');
  final raf = file.openSync(mode: FileMode.write);
  final bytes = [65, 66, 67, 68]; // ABCD
  raf.writeFromSync(bytes, 1, 3); // write B, C
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["BC"]
    );
}

#[test]
fn random_access_file_read_byte_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_read_byte.bin');
  file.writeAsBytesSync([100, 200]);
  final raf = file.openSync(mode: FileMode.read);
  print(raf.readByteSync());
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["100"]
    );
}

#[test]
fn random_access_file_read_sync_buffer() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_read_sync.bin');
  file.writeAsBytesSync([10, 20, 30, 40]);
  final raf = file.openSync(mode: FileMode.read);
  final buffer = raf.readSync(3);
  print(buffer.length);
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn random_access_file_read_into_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
import 'dart:typed_data';
void main() {
  final file = File('test_raf_read_into.bin');
  file.writeAsBytesSync([1, 2, 3, 4]);
  final raf = file.openSync(mode: FileMode.read);
  final buffer = Uint8List(5);
  final bytesRead = raf.readIntoSync(buffer, 1, 4);
  print('$bytesRead:${buffer[1]}');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["3:1"]
    );
}

#[test]
fn random_access_file_set_position_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_seek.bin');
  file.writeAsBytesSync([10, 20, 30, 40]);
  final raf = file.openSync(mode: FileMode.read);
  raf.setPositionSync(2);
  print(raf.readByteSync());
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["30"]
    );
}

#[test]
fn random_access_file_get_position_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_tell.bin');
  file.writeAsBytesSync([1, 2, 3]);
  final raf = file.openSync(mode: FileMode.read);
  raf.readByteSync();
  print(raf.positionSync());
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn random_access_file_length_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_len.bin');
  file.writeAsBytesSync([1, 2, 3, 4, 5]);
  final raf = file.openSync(mode: FileMode.read);
  print(raf.lengthSync());
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn random_access_file_truncate_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_trunc.bin');
  file.writeAsStringSync('truncate_me');
  final raf = file.openSync(mode: FileMode.write);
  raf.truncateSync(4);
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["trun"]
    );
}

#[test]
fn random_access_file_truncate_sync_expand() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_trunc_exp.bin');
  file.writeAsStringSync('a');
  final raf = file.openSync(mode: FileMode.write);
  raf.truncateSync(3);
  raf.closeSync();
  print(file.lengthSync());
  file.deleteSync();
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn random_access_file_flush_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_flush.bin');
  final raf = file.openSync(mode: FileMode.write);
  raf.writeStringSync('data');
  raf.flushSync();
  raf.closeSync();
  print(file.readAsStringSync());
  file.deleteSync();
}
"#
        ),
        vec!["data"]
    );
}

#[test]
fn random_access_file_lock_sync_exclusive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_lock_ex.bin');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  print('locked');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["locked"]
    );
}

#[test]
fn random_access_file_lock_sync_shared() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_lock_sh.bin');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.read);
  raf.lockSync(FileLock.shared);
  print('shared_lock');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["shared_lock"]
    );
}

#[test]
fn random_access_file_unlock_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_unlock.bin');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  raf.unlockSync();
  print('unlocked');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["unlocked"]
    );
}

#[test]
fn random_access_file_operations_after_close_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_close_err.bin');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.read);
  raf.closeSync();
  try {
    raf.readByteSync();
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    file.deleteSync();
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn random_access_file_path_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_path.bin');
  final raf = file.openSync(mode: FileMode.write);
  print(raf.path.contains('test_raf_path'));
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn random_access_file_write_string_sync_encoding() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
import 'dart:convert';
void main() {
  final file = File('test_raf_enc.bin');
  final raf = file.openSync(mode: FileMode.write);
  raf.writeStringSync('Därt', encoding: utf8);
  raf.closeSync();
  print(file.readAsStringSync(encoding: utf8));
  file.deleteSync();
}
"#
        ),
        vec!["Därt"]
    );
}

#[test]
fn random_access_file_seek_beyond_end() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_seek_end.bin');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.setPositionSync(10);
  raf.writeStringSync('end');
  raf.closeSync();
  print(file.lengthSync());
  file.deleteSync();
}
"#
        ),
        vec!["13"] // 10 + 3
    );
}

#[test]
fn random_access_file_read_into_sync_out_of_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_raf_oob.bin');
  file.writeAsBytesSync([1, 2]);
  final raf = file.openSync(mode: FileMode.read);
  try {
    raf.readIntoSync([0, 0, 0], 1, 5); // 5 > 3
  } on RangeError {
    print('RangeError thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["RangeError thrown"]
    );
}
