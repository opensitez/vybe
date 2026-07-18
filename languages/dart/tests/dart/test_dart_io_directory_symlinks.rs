use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io Link & Symlinks
// ═══════════════════════════════════════════════════════════

#[test]
fn link_create_sync_to_file() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('test_link_tgt.txt');
  file.writeAsStringSync('data');
  final link = Link('test_link_src.lnk');
  link.createSync(file.path);
  print(link.existsSync());
  link.deleteSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_create_sync_to_directory() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final dir = Directory('test_link_dir');
  dir.createSync();
  final link = Link('test_link_dir.lnk');
  link.createSync(dir.path);
  print(link.existsSync());
  link.deleteSync();
  dir.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_create_sync_recursive_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('non_existent_dir/my_link.lnk');
  try {
    link.createSync('target.txt', recursive: false);
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
fn link_create_sync_recursive_true() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('test_link_rec/my_link.lnk');
  link.createSync('target.txt', recursive: true);
  print(link.existsSync());
  Directory('test_link_rec').deleteSync(recursive: true);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_target_sync_valid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('target_file.txt');
  file.createSync();
  final link = Link('link_tgt_sync.lnk');
  link.createSync(file.path);
  print(link.targetSync() == file.path);
  link.deleteSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_target_sync_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('does_not_exist_link.lnk');
  try {
    link.targetSync();
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
fn link_target_sync_not_a_link_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('not_a_link.txt');
  file.createSync();
  final link = Link(file.path); // Point Link object to an actual File
  try {
    link.targetSync();
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
fn link_delete_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('to_be_deleted.lnk');
  link.createSync('dummy_target.txt');
  link.deleteSync();
  print(link.existsSync());
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn link_delete_sync_leaves_target_intact() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('intact_target.txt');
  file.createSync();
  final link = Link('del_link.lnk');
  link.createSync(file.path);
  link.deleteSync();
  print(file.existsSync());
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_rename_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('old_link_name.lnk');
  link.createSync('target.txt');
  final renamed = link.renameSync('new_link_name.lnk');
  print('${link.existsSync()}:${renamed.existsSync()}');
  renamed.deleteSync();
}
"#
        ),
        vec!["false:true"]
    );
}

#[test]
fn link_rename_sync_changes_path() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('link1.lnk');
  link.createSync('tgt');
  final renamed = link.renameSync('link2.lnk');
  print(renamed.path);
  renamed.deleteSync();
}
"#
        ),
        vec!["link2.lnk"]
    );
}

#[test]
fn link_rename_to_existing_file_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('link_to_rename.lnk');
  link.createSync('tgt');
  final file = File('existing_file_for_link.txt');
  file.createSync();
  try {
    link.renameSync(file.path);
  } on FileSystemException {
    print('FileSystemException thrown');
  } finally {
    link.deleteSync();
    file.deleteSync();
  }
}
"#
        ),
        vec!["FileSystemException thrown"]
    );
}

#[test]
fn link_update_sync_changes_target() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('update_link.lnk');
  link.createSync('tgt1.txt');
  link.updateSync('tgt2.txt');
  print(link.targetSync());
  link.deleteSync();
}
"#
        ),
        vec!["tgt2.txt"]
    );
}

#[test]
fn link_update_sync_non_existent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('no_update_link.lnk');
  try {
    link.updateSync('tgt.txt');
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
fn link_absolute_getter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('rel_link.lnk');
  print(link.absolute.isAbsolute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_stat_sync_without_following() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('stat_link.lnk');
  link.createSync('non_existent_target.txt');
  // FileStat.statSync(path) follows links by default unless told not to,
  // but link.statSync doesn't follow link? Wait, Link doesn't have statSync natively in Dart.
  // It has FileStat.statSync(path).
  final stat = FileStat.statSync(link.path);
  // If followed, it's notFound. If not followed (not default), it's link.
  print(stat.type == FileSystemEntityType.notFound);
  link.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_type_sync_follows_by_default() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('link_type_target.txt');
  file.createSync();
  final link = Link('link_type.lnk');
  link.createSync(file.path);
  final type = FileSystemEntity.typeSync(link.path);
  print(type == FileSystemEntityType.file);
  link.deleteSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_type_sync_no_follow() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link = Link('link_type_no_follow.lnk');
  link.createSync('tgt');
  final type = FileSystemEntity.typeSync(link.path, followLinks: false);
  print(type == FileSystemEntityType.link);
  link.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_create_cyclic_chain() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final link1 = Link('l1.lnk');
  final link2 = Link('l2.lnk');
  link1.createSync('l2.lnk');
  link2.createSync('l1.lnk');
  // targetSync just reads the link, it doesn't resolve it. So it won't infinite loop.
  print(link1.targetSync() == 'l2.lnk');
  link1.deleteSync();
  link2.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn link_resolve_symbolic_links_sync_on_link() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final file = File('resolve_target.txt');
  file.createSync();
  final link = Link('resolve_link.lnk');
  link.createSync(file.path);
  final resolved = link.resolveSymbolicLinksSync();
  print(resolved.isNotEmpty);
  link.deleteSync();
  file.deleteSync();
}
"#
        ),
        vec!["true"]
    );
}
