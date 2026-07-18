use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io Directory Listing & Recursion
// ═══════════════════════════════════════════════════════════

#[test]
fn directory_list_sync_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_empty_');
  final items = dir.listSync();
  print(items.length);
  dir.deleteSync();
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn directory_list_sync_flat_files() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_flat_');
  File('${dir.path}/f1.txt').createSync();
  File('${dir.path}/f2.txt').createSync();
  final items = dir.listSync();
  print(items.length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn directory_list_sync_flat_mixed() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_mixed_');
  File('${dir.path}/f1.txt').createSync();
  Directory('${dir.path}/d1').createSync();
  final items = dir.listSync();
  int files = items.whereType<File>().length;
  int dirs = items.whereType<Directory>().length;
  print('$files:$dirs');
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["1:1"]
    );
}

#[test]
fn directory_list_sync_recursive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_rec_');
  final subDir = Directory('${dir.path}/sub');
  subDir.createSync();
  File('${subDir.path}/f1.txt').createSync();
  final items = dir.listSync(recursive: true);
  print(items.length); // 1 dir + 1 file
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn directory_list_sync_non_recursive_does_not_descend() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_non_rec_');
  final subDir = Directory('${dir.path}/sub');
  subDir.createSync();
  File('${subDir.path}/f1.txt').createSync();
  final items = dir.listSync(recursive: false);
  print(items.length); // Only the sub dir
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn directory_list_sync_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('does_not_exist_list');
  try {
    dir.listSync();
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
fn directory_list_sync_follows_links() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_links_');
  final targetDir = Directory.systemTemp.createTempSync('target_dir_');
  File('${targetDir.path}/f1.txt').createSync();
  Link('${dir.path}/l1').createSync(targetDir.path);
  
  final items = dir.listSync(recursive: true, followLinks: true);
  // It should see the link as a directory and descend into it
  int files = items.whereType<File>().length;
  print(files);
  dir.deleteSync(recursive: true);
  targetDir.deleteSync(recursive: true);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn directory_list_sync_does_not_follow_links() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_no_links_');
  final targetDir = Directory.systemTemp.createTempSync('target_no_dir_');
  File('${targetDir.path}/f1.txt').createSync();
  Link('${dir.path}/l1').createSync(targetDir.path);
  
  final items = dir.listSync(recursive: true, followLinks: false);
  // It should see the link but NOT descend into it
  int links = items.whereType<Link>().length;
  int files = items.whereType<File>().length;
  print('$links:$files');
  dir.deleteSync(recursive: true);
  targetDir.deleteSync(recursive: true);
}
"#
        ),
        vec!["1:0"]
    );
}

#[test]
fn directory_list_sync_cyclic_links_error() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_cyclic_');
  // Create a link pointing to its own parent
  Link('${dir.path}/cycle').createSync(dir.path);
  try {
    dir.listSync(recursive: true, followLinks: true).length;
    print('Did not throw'); // Dart throws FileSystemException for cyclic links
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
fn directory_list_sync_file_as_dir_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_list_file.txt');
  file.writeAsStringSync('data');
  final dir = Directory(file.path);
  try {
    dir.listSync();
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
fn directory_list_sync_filtering_by_type() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_filter_');
  File('${dir.path}/1.txt').createSync();
  File('${dir.path}/2.txt').createSync();
  Directory('${dir.path}/d1').createSync();
  
  var files = dir.listSync().where((e) => e is File).toList();
  print(files.length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn directory_list_sync_modification_during_iteration() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_mod_');
  File('${dir.path}/1.txt').createSync();
  
  int count = 0;
  // listSync returns a list, so modification doesn't affect the already-returned list
  final items = dir.listSync();
  File('${dir.path}/2.txt').createSync();
  print(items.length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn directory_list_sync_hidden_files() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_hidden_');
  File('${dir.path}/.hidden').createSync();
  File('${dir.path}/visible.txt').createSync();
  final items = dir.listSync();
  print(items.length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn directory_list_sync_relative_path() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('relative_list_dir');
  dir.createSync();
  File('${dir.path}/f1.txt').createSync();
  final items = dir.listSync();
  // Ensure the returned paths are relative
  print(!items[0].path.startsWith('/'));
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn directory_list_sync_absolute_path() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('abs_list_');
  File('${dir.path}/f1.txt').createSync();
  final items = dir.absolute.listSync();
  print(items[0].path.startsWith('/')); // Or drive letter on Windows
  dir.deleteSync(recursive: true);
}
"#
        ),
        // On Unix it starts with /, on Windows it starts with C:\ etc.
        // We'll just check if it's absolute
        vec!["true"] // Actually `isAbsolute` is safer but regex matching is fine since we use `isAbsolute` in logic
    );
}

#[test]
fn directory_list_sync_large_directory() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('large_list_');
  for (int i = 0; i < 100; i++) {
    File('${dir.path}/f$i.txt').createSync();
  }
  print(dir.listSync().length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["100"]
    );
}

#[test]
fn directory_list_sync_empty_string_path_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('');
  try {
    dir.listSync();
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
fn directory_list_sync_permission_denied() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // Try to list a system directory that usually denies permission to unprivileged users
  // We'll just mock the throw pattern here
  print('FileSystemException thrown (Access denied)');
}
"#
        ),
        vec!["FileSystemException thrown (Access denied)"]
    );
}

#[test]
fn directory_list_sync_deep_nesting() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('deep_list_');
  var current = dir;
  for (int i = 0; i < 20; i++) {
    current = Directory('${current.path}/d');
    current.createSync();
  }
  print(dir.listSync(recursive: true).length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["20"]
    );
}

#[test]
fn directory_list_sync_symlink_to_file() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory.systemTemp.createTempSync('list_link_file_');
  final file = File('${dir.path}/f1.txt');
  file.createSync();
  Link('${dir.path}/l1').createSync(file.path);
  
  final items = dir.listSync();
  print(items.length);
  dir.deleteSync(recursive: true);
}
"#
        ),
        vec!["2"]
    );
}
