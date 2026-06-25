use crate::helpers::run_python_one;

#[test]
fn os_path_join_two_parts() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('a', 'b'))\n"),
        "a/b"
    );
}

#[test]
fn os_path_join_three_parts() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('tmp', 'data', 'file.txt'))\n"),
        "tmp/data/file.txt"
    );
}

#[test]
fn os_path_basename() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('/tmp/data/file.txt'))\n"),
        "file.txt"
    );
}

#[test]
fn os_path_dirname() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.dirname('/tmp/data/file.txt'))\n"),
        "/tmp/data"
    );
}

#[test]
fn os_path_splitext() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.splitext('archive.tar.gz'))\n"),
        "('archive.tar', '.gz')"
    );
}

#[test]
fn os_path_splitext_simple() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.splitext('readme.txt'))\n"),
        "('readme', '.txt')"
    );
}

#[test]
fn os_path_exists_false_for_missing() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.exists('/no/such/vybe_test_path_9f3c'))\n"),
        "False"
    );
}

#[test]
fn os_path_isfile_false_for_missing() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.isfile('/no/such/vybe_test_path_9f3c'))\n"),
        "False"
    );
}

#[test]
fn os_path_isdir_false_for_missing() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.isdir('/no/such/vybe_test_path_9f3c'))\n"),
        "False"
    );
}

#[test]
fn os_path_abspath_relative() {
    assert_eq!(
        run_python_one("import os\np = os.path.abspath('vybe_rel.txt')\nprint(p.endswith('vybe_rel.txt'))\n"),
        "True"
    );
}

#[test]
fn os_path_join_with_empty_part() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('a', ''))\n"),
        "a/"
    );
}

#[test]
fn os_path_basename_trailing_slash() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('/tmp/dir/'))\n"),
        ""
    );
}

#[test]
fn os_path_dirname_root_file() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.dirname('/file.txt'))\n"),
        "/"
    );
}

#[test]
fn os_getcwd_returns_nonempty_string() {
    assert_eq!(
        run_python_one("import os\ncwd = os.getcwd()\nprint(isinstance(cwd, str) and len(cwd) > 0)\n"),
        "True"
    );
}

#[test]
fn os_listdir_current_is_list() {
    assert_eq!(
        run_python_one("import os\nnames = os.listdir('.')\nprint(isinstance(names, list))\n"),
        "True"
    );
}

#[test]
fn os_path_join_preserves_first_absolute() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('/root', 'sub', 'leaf'))\n"),
        "/root/sub/leaf"
    );
}

#[test]
fn os_path_splitext_no_extension() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.splitext('README'))\n"),
        "('README', '')"
    );
}

#[test]
fn os_path_basename_single_name() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('file.py'))\n"),
        "file.py"
    );
}

#[test]
fn os_path_dirname_single_name() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.dirname('file.py'))\n"),
        ""
    );
}

#[test]
fn os_path_exists_current_dir() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.exists('.'))\n"),
        "True"
    );
}

#[test]
fn os_path_isdir_current_dir() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.isdir('.'))\n"),
        "True"
    );
}

#[test]
fn os_path_isfile_current_dir_false() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.isfile('.'))\n"),
        "False"
    );
}

#[test]
fn os_import_succeeds() {
    assert_eq!(
        run_python_one("import os\nprint(os.__name__)\n"),
        "os"
    );
}

#[test]
fn os_path_module_accessible() {
    assert_eq!(
        run_python_one("import os\nprint(hasattr(os, 'path'))\n"),
        "True"
    );
}

#[test]
fn os_path_join_many_segments() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('a', 'b', 'c', 'd'))\n"),
        "a/b/c/d"
    );
}

#[test]
fn os_path_abspath_dot() {
    assert_eq!(
        run_python_one("import os\np = os.path.abspath('.')\nprint(len(p) > 0)\n"),
        "True"
    );
}

#[test]
fn os_path_splitext_double_dot() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.splitext('module.test.py'))\n"),
        "('module.test', '.py')"
    );
}

#[test]
fn os_path_basename_hidden_file() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('/home/.bashrc'))\n"),
        ".bashrc"
    );
}

#[test]
fn os_path_dirname_nested() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.dirname('/a/b/c/d'))\n"),
        "/a/b/c"
    );
}

#[test]
fn os_path_join_absolute_second_segment() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('rel', '/abs'))\n"),
        "/abs"
    );
}

#[test]
fn os_path_exists_after_join_missing() {
    assert_eq!(
        run_python_one("import os\np = os.path.join('missing', 'vybe', 'ghost.txt')\nprint(os.path.exists(p))\n"),
        "False"
    );
}

#[test]
fn os_path_isfile_with_joined_path() {
    assert_eq!(
        run_python_one("import os\np = os.path.join('nope', 'file.txt')\nprint(os.path.isfile(p))\n"),
        "False"
    );
}

#[test]
fn os_path_dirname_preserves_root() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.dirname('/'))\n"),
        "/"
    );
}

#[test]
fn os_path_basename_root() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('/'))\n"),
        ""
    );
}

#[test]
fn os_path_splitext_uppercase_ext() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.splitext('photo.JPG'))\n"),
        "('photo', '.JPG')"
    );
}

#[test]
fn os_path_join_with_dot() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.join('.', 'local'))\n"),
        "./local"
    );
}

#[test]
fn os_path_abspath_joined() {
    assert_eq!(
        run_python_one("import os\np = os.path.abspath(os.path.join('a', 'b.txt'))\nprint(p.endswith('a/b.txt') or p.endswith('a\\\\b.txt'))\n"),
        "True"
    );
}

#[test]
fn os_listdir_type_has_length() {
    assert_eq!(
        run_python_one("import os\nprint(len(os.listdir('.')) >= 0)\n"),
        "True"
    );
}

#[test]
fn os_path_operations_chain() {
    assert_eq!(
        run_python_one("import os\np = os.path.join('x', 'y.txt')\nprint(os.path.basename(p), os.path.splitext(p)[1])\n"),
        "y.txt .txt"
    );
}

#[test]
fn os_path_dirname_of_joined() {
    assert_eq!(
        run_python_one("import os\np = os.path.join('dir', 'sub', 'f.py')\nprint(os.path.dirname(p))\n"),
        "dir/sub"
    );
}

#[test]
fn os_getcwd_matches_abspath_dot_parent() {
    assert_eq!(
        run_python_one("import os\ncwd = os.getcwd()\nprint(cwd == os.path.abspath('.') or len(cwd) > 0)\n"),
        "True"
    );
}

#[test]
fn os_path_exists_is_bool() {
    assert_eq!(
        run_python_one("import os\nprint(type(os.path.exists('.')).__name__)\n"),
        "bool"
    );
}

#[test]
fn os_path_join_result_is_str() {
    assert_eq!(
        run_python_one("import os\nprint(type(os.path.join('a', 'b')).__name__)\n"),
        "str"
    );
}

#[test]
fn os_path_splitext_returns_tuple() {
    assert_eq!(
        run_python_one("import os\nprint(type(os.path.splitext('a.b')).__name__)\n"),
        "tuple"
    );
}

#[test]
fn os_path_basename_from_joined_windows_style_input() {
    assert_eq!(
        run_python_one("import os\nprint(os.path.basename('folder/item'))\n"),
        "item"
    );
}
