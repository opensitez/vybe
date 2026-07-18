use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io Directory Creation/Deletion
// ═══════════════════════════════════════════════════════════

#[test]
fn directory_create_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_create_sync_dir');
  dir.createSync();
  print(dir.existsSync());
  dir.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_create_nested_without_recursive_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_nested_1/test_nested_2');
  try {
    dir.createSync();
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
fn directory_create_nested_with_recursive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_recursive_1/test_recursive_2');
  dir.createSync(recursive: true);
  print(dir.existsSync());
  Directory('test_recursive_1').deleteSync(recursive: true);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_delete_empty_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_empty_delete');
  dir.createSync();
  dir.deleteSync();
  print(dir.existsSync());
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn directory_delete_non_empty_without_recursive_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_non_empty_del');
  dir.createSync();
  File('${dir.path}/temp.txt').writeAsStringSync('data');
  try {
    dir.deleteSync();
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    dir.deleteSync(recursive: true);
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn directory_delete_non_empty_with_recursive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_non_empty_rec_del');
  dir.createSync();
  File('${dir.path}/temp.txt').writeAsStringSync('data');
  dir.deleteSync(recursive: true);
  print(dir.existsSync());
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn directory_delete_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('does_not_exist_dir');
  try {
    dir.deleteSync();
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
fn directory_exists_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_exists_sync');
  print(dir.existsSync());
  dir.createSync();
  print(dir.existsSync());
  dir.deleteSync();
}
"#
        ),
        vec!["false\ntrue"]
    );
}

#[test]
fn directory_rename_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir1 = Directory('test_rename_src');
  final dir2 = Directory('test_rename_dest');
  dir1.createSync();
  final renamed = dir1.renameSync(dir2.path);
  print('${dir1.existsSync()}:${renamed.existsSync()}');
  renamed.deleteSync();
}
"#
        ),
        vec!["false:true"]
    );
}

#[test]
fn directory_rename_to_existing_file_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_rename_src_2');
  dir.createSync();
  final file = File('test_rename_dest_file.txt');
  file.writeAsStringSync('data');
  try {
    dir.renameSync(file.path);
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    dir.deleteSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn directory_create_temp_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync();
  print(dir.existsSync());
  dir.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_create_temp_sync_with_prefix() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('my_prefix_');
  print(dir.path.contains('my_prefix_'));
  dir.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_current_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final current = Directory.current;
  print(current.isAbsolute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_current_setter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final original = Directory.current;
  final temp = Directory.systemTemp.createTempSync();
  Directory.current = temp;
  print(Directory.current.path == temp.path);
  Directory.current = original;
  temp.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_system_temp_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final temp = Directory.systemTemp;
  print(temp.existsSync());
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_absolute_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('some_relative_dir');
  print(dir.absolute.isAbsolute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_resolve_symbolic_links_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final current = Directory.current;
  final resolved = current.resolveSymbolicLinksSync();
  print(resolved.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_create_invalid_characters_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('invalid\x00dir');
  try {
    dir.createSync();
  } on FileSystemException {
    print('FileSystemException thrown');
  } catch (e) {
    print('ArgumentError thrown'); // Depending on platform, may throw ArgumentError
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn directory_delete_root_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('/');
  try {
    dir.deleteSync();
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
fn directory_stat_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_stat_dir');
  dir.createSync();
  final stat = dir.statSync();
  print(stat.type == FileSystemEntityType.directory);
  dir.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}
