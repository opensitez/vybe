use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io File Stat & Metadata
// ═══════════════════════════════════════════════════════════

#[test]
fn file_stat_sync_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_basic.txt');
  file.writeAsStringSync('stat test');
  final stat = file.statSync();
  print(stat.type == FileSystemEntityType.file);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_non_existent() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final stat = File('does_not_exist_stat.txt').statSync();
  print(stat.type == FileSystemEntityType.notFound);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_size() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_size.txt');
  file.writeAsStringSync('12345');
  final stat = file.statSync();
  print(stat.size);
  file.deleteSync();
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn file_stat_sync_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_mode.txt');
  file.writeAsStringSync('mode');
  final stat = file.statSync();
  // We can't guarantee the exact mode across OSes, just that it's > 0
  print(stat.mode > 0);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_modified_time() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_modified.txt');
  file.writeAsStringSync('mod');
  final stat = file.statSync();
  print(stat.modified.year >= 2024);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_accessed_time() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_accessed.txt');
  file.writeAsStringSync('acc');
  final stat = file.statSync();
  print(stat.accessed.year >= 2024);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_changed_time() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_stat_changed.txt');
  file.writeAsStringSync('chg');
  final stat = file.statSync();
  print(stat.changed.year >= 2024);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_last_modified_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_last_mod.txt');
  file.writeAsStringSync('test');
  final modified = file.lastModifiedSync();
  print(modified.year >= 2024);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_set_last_modified_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_set_last_mod.txt');
  file.writeAsStringSync('test');
  final target = DateTime(2030, 1, 1);
  file.setLastModifiedSync(target);
  final actual = file.lastModifiedSync();
  print('${actual.year}:${actual.month}:${actual.day}');
  file.deleteSync();
}
"#
        ),
        vec!["2030:1:1"]
    );
}

#[test]
fn file_set_last_accessed_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_set_last_acc.txt');
  file.writeAsStringSync('test');
  final target = DateTime(2031, 2, 2);
  file.setLastAccessedSync(target);
  final stat = file.statSync();
  print('${stat.accessed.year}:${stat.accessed.month}:${stat.accessed.day}');
  file.deleteSync();
}
"#
        ),
        vec!["2031:2:2"]
    );
}

#[test]
fn file_length_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_len.txt');
  file.writeAsBytesSync([1, 2, 3, 4, 5, 6, 7]);
  print(file.lengthSync());
  file.deleteSync();
}
"#
        ),
        vec!["7"]
    );
}

#[test]
fn file_length_sync_non_existent() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('does_not_exist_len.txt');
  try {
    file.lengthSync();
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
fn file_absolute_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('relative.txt');
  final abs = file.absolute;
  print(abs.isAbsolute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_parent_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('dir/file.txt');
  final parent = file.parent;
  print(parent.path.endsWith('dir'));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_resolve_symbolic_links_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_sym_resolve.txt');
  file.writeAsStringSync('data');
  final resolved = file.resolveSymbolicLinksSync();
  print(resolved.isNotEmpty);
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_stat_sync_on_symlink() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_sym_stat.txt');
  file.writeAsStringSync('data');
  final link = Link('test_sym_link.txt');
  link.createSync(file.path);
  final stat = File(link.path).statSync(); // follows link by default
  print(stat.type == FileSystemEntityType.file);
  link.deleteSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn file_rename_sync_changes_metadata() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_rename_meta.txt');
  file.writeAsStringSync('meta');
  final renamed = file.renameSync('test_renamed_meta.txt');
  print(renamed.statSync().size);
  renamed.deleteSync();
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn file_copy_sync_preserves_size() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_copy_meta.txt');
  file.writeAsStringSync('copy data');
  final copied = file.copySync('test_copied_meta.txt');
  print(copied.statSync().size);
  file.deleteSync();
  copied.deleteSync();
}
"#
        ),
        vec!["9"]
    );
}

#[test]
fn file_copy_sync_to_non_existent_dir_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_copy_err.txt');
  file.writeAsStringSync('err');
  try {
    file.copySync('does_not_exist_dir/copied.txt');
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
fn file_set_last_modified_sync_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('does_not_exist_mod.txt');
  try {
    file.setLastModifiedSync(DateTime.now());
  } on FileSystemException {
    print('FileSystemException thrown');
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}
