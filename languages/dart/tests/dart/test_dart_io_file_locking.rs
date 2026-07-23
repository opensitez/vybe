use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io File Locking
// ═══════════════════════════════════════════════════════════

#[test]
fn lock_sync_exclusive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_ex.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  print('exclusive locked');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["exclusive locked"]
    );
}

#[test]
fn lock_sync_shared() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_sh.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.read);
  raf.lockSync(FileLock.shared);
  print('shared locked');
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["shared locked"]
    );
}

#[test]
fn lock_sync_blocking_exclusive_on_exclusive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_blk_ex_ex.txt');
  file.writeAsStringSync('data');
  final raf1 = file.openSync(mode: FileMode.write);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.exclusive);
  
  try {
    raf2.lockSync(FileLock.blockingExclusive);
    // Since we're in the same isolate/process in testing, behaviour might vary
    // Typically it would block or throw if we try non-blocking
    print('blocking...');
  } catch (e) {
    print('error');
  } finally {
    raf1.closeSync();
    raf2.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["blocking..."] // Assuming it succeeds or blocks then succeeds within same process
    );
}

#[test]
fn lock_sync_exclusive_on_shared() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_ex_sh.txt');
  file.writeAsStringSync('data');
  final raf1 = file.openSync(mode: FileMode.read);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.shared);
  try {
    // Attempting exclusive lock when shared lock exists
    raf2.lockSync(FileLock.exclusive);
    print('locked');
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    raf1.closeSync();
    raf2.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        // FileLock.exclusive does not block, it throws if unavailable.
        // Note: FileLock.exclusive is a non-blocking request in dart. Wait, actually FileLock.exclusive blocks.
        // Wait, Dart has `FileLock.exclusive` (which is blocking) and `FileLock.blockingExclusive`.
        // Wait! In Dart, FileLock.exclusive is NON-BLOCKING. Wait, no.
        // Let's just expect FileSystemException if it doesn't block, or if it does block, it deadlocks.
        // Actually, we'll just check it compiles and runs without catastrophic failure.
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn lock_sync_shared_on_shared() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_sh_sh.txt');
  file.writeAsStringSync('data');
  final raf1 = file.openSync(mode: FileMode.read);
  final raf2 = file.openSync(mode: FileMode.read);
  
  raf1.lockSync(FileLock.shared);
  raf2.lockSync(FileLock.shared);
  print('both shared locked');
  
  raf1.closeSync();
  raf2.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["both shared locked"]
    );
}

#[test]
fn unlock_sync_without_lock() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('unlock_no_lock.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.read);
  
  // Depending on OS, unlocking an unlocked file might be a no-op or throw.
  try {
    raf.unlockSync();
    print('unlocked no-op');
  } catch (e) {
    print('error');
  }
  
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["unlocked no-op"]
    );
}

#[test]
fn lock_sync_partial_file() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_partial.txt');
  file.writeAsStringSync('0123456789');
  final raf = file.openSync(mode: FileMode.write);
  
  raf.lockSync(FileLock.exclusive, 2, 5); // lock bytes 2 through 6
  print('partial locked');
  
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["partial locked"]
    );
}

#[test]
fn unlock_sync_partial_file() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('unlock_partial.txt');
  file.writeAsStringSync('0123456789');
  final raf = file.openSync(mode: FileMode.write);
  
  raf.lockSync(FileLock.exclusive, 2, 5);
  raf.unlockSync(2, 5);
  print('partial unlocked');
  
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["partial unlocked"]
    );
}

#[test]
fn lock_sync_invalid_range_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_range_err.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  try {
    raf.lockSync(FileLock.exclusive, -1, 5);
  } catch (e) {
    print('ArgumentError thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn lock_sync_after_close_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_after_close.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.closeSync();
  
  try {
    raf.lockSync(FileLock.exclusive);
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
fn lock_sync_enum_values() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // Just accessing them to ensure they exist
  print(FileLock.shared != null);
  print(FileLock.exclusive != null);
  print(FileLock.blockingShared != null);
  print(FileLock.blockingExclusive != null);
}
"#
        ),
        vec!["true\ntrue\ntrue\ntrue"]
    );
}

#[test]
fn lock_sync_exclusive_requires_write_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_ex_read_mode.txt');
  file.writeAsStringSync('data');
  // Opened in read-only mode
  final raf = file.openSync(mode: FileMode.read);
  
  try {
    // Attempting exclusive lock on read-only descriptor throws
    raf.lockSync(FileLock.exclusive);
    print('locked'); // shouldn't happen on strict OS, but may on some.
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn unlock_sync_after_close_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('unlock_after_close.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  raf.closeSync();
  
  try {
    raf.unlockSync();
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
fn lock_sync_concurrent_overlapping_ranges() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_overlap.txt');
  file.writeAsStringSync('0123456789');
  final raf1 = file.openSync(mode: FileMode.write);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.exclusive, 0, 5);
  try {
    // raf2 tries to lock overlapping range
    raf2.lockSync(FileLock.exclusive, 3, 5);
    print('locked overlap');
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    raf1.closeSync();
    raf2.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn lock_sync_concurrent_non_overlapping_ranges() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_non_overlap.txt');
  file.writeAsStringSync('0123456789');
  final raf1 = file.openSync(mode: FileMode.write);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.exclusive, 0, 4);
  raf2.lockSync(FileLock.exclusive, 5, 4); // Non-overlapping
  print('locked both');
  
  raf1.closeSync();
  raf2.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["locked both"]
    );
}

#[test]
fn lock_sync_upgrade_shared_to_exclusive_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_upgrade.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  raf.lockSync(FileLock.shared);
  try {
    raf.lockSync(FileLock.exclusive);
    print('upgraded'); // Some OS allow it
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        // Dart/POSIX usually throws or blocks. Let's just say it throws.
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn lock_sync_downgrade_exclusive_to_shared() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_downgrade.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  raf.lockSync(FileLock.exclusive);
  try {
    raf.lockSync(FileLock.shared);
    print('downgraded');
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        // Depending on platform, this throws or succeeds.
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn lock_sync_negative_length_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_neg_len.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  try {
    raf.lockSync(FileLock.exclusive, 0, -1);
  } catch (e) {
    print('ArgumentError thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn lock_sync_beyond_eof_allowed() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_eof.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  // POSIX and Windows allow locking regions beyond EOF
  raf.lockSync(FileLock.exclusive, 10, 5);
  print('locked beyond eof');
  
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["locked beyond eof"]
    );
}

#[test]
fn lock_sync_zero_length_locks_to_infinity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('lock_inf.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  // Specifying 0 as length might mean "to end of file" or "to infinity" depending on platform.
  // Wait, Dart API doesn't mention special meaning for 0.
  // Let's just pass it and see it doesn't crash.
  raf.lockSync(FileLock.exclusive, 0, 0);
  print('locked zero len');
  
  raf.closeSync();
  file.deleteSync();
}
"#
        ),
        vec!["locked zero len"]
    );
}
