use super::helpers::run_python;

// site — getsitepackages, getuserbase, getusersitepackages, PREFIXES, ENABLE_USER_SITE, USER_BASE, USER_SITE, addsitedir, addpackage, makepath, pth processing

#[test]
fn test_site_getsitepackages_contains_expected_directory_structure() {
    let out = run_python(
        r#"
import site, sys
paths = site.getsitepackages()
print(any(sys.prefix in p or sys.exec_prefix in p for p in paths))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_getuserbase_returns_non_empty_path_str() {
    let out = run_python(
        r#"
import site
base = site.getuserbase()
print(base.startswith('/') or len(base) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_getusersitepackages_starts_with_user_base() {
    let out = run_python(
        r#"
import site
base = site.getuserbase()
user_site = site.getusersitepackages()
print(user_site.startswith(base))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_prefixes_list_contains_sys_prefix() {
    let out = run_python(
        r#"
import site, sys
print(sys.prefix in site.PREFIXES)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_user_base_matches_getuserbase() {
    let out = run_python(
        r#"
import site
print(site.USER_BASE == site.getuserbase())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_user_site_matches_getusersitepackages() {
    let out = run_python(
        r#"
import site
print(site.USER_SITE == site.getusersitepackages())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_addsitedir_adds_to_sys_path() {
    let out = run_python(
        r#"
import site, sys, tempfile
with tempfile.TemporaryDirectory() as tmpdir:
    site.addsitedir(tmpdir)
    print(tmpdir in sys.path)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_addsitedir_processes_pth_file() {
    let out = run_python(
        r#"
import site, sys, tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    sub_dir = os.path.join(tmpdir, 'added_via_pth')
    os.makedirs(sub_dir)
    pth_file = os.path.join(tmpdir, 'package.pth')
    with open(pth_file, 'w') as f:
        f.write('added_via_pth\n')
    site.addsitedir(tmpdir)
    print(sub_dir in sys.path)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_addsitedir_pth_file_import_exec() {
    let out = run_python(
        r#"
import site, sys, tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    pth_file = os.path.join(tmpdir, 'exec_code.pth')
    with open(pth_file, 'w') as f:
        f.write("import sys; sys._pth_test_var = 'executed'\n")
    site.addsitedir(tmpdir)
    print(getattr(sys, '_pth_test_var', None))
"#,
    );
    assert_eq!(out, vec!["executed"]);
}

#[test]
fn test_site_getsitepackages_with_custom_prefixes() {
    let out = run_python(
        r#"
import site
prefixes = ['/custom/prefix1', '/custom/prefix2']
res = site.getsitepackages(prefixes)
print(any('/custom/prefix1' in p for p in res))
print(any('/custom/prefix2' in p for p in res))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_site_addsitedir_duplicate_prevention() {
    let out = run_python(
        r#"
import site, sys, tempfile
with tempfile.TemporaryDirectory() as tmpdir:
    site.addsitedir(tmpdir)
    count1 = sys.path.count(tmpdir)
    site.addsitedir(tmpdir)
    count2 = sys.path.count(tmpdir)
print(count1 == 1)
print(count2 == 1)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_site_set_quit_helper_creates_quit_and_exit() {
    let out = run_python(
        r#"
import site, builtins
site.setquit()
q = str(builtins.quit)
print('Use quit()' in q or 'exit' in q)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_set_copyright_helper() {
    let out = run_python(
        r#"
import site, builtins
site.setcopyright()
c = str(builtins.copyright)
print('Copyright' in c)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_set_helper_creates_help() {
    let out = run_python(
        r#"
import site, builtins
site.sethelper()
h = str(builtins.help)
print('Type help()' in h or 'interactive help' in h)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_makepath_helper_joins_and_normalizes() {
    let out = run_python(
        r#"
import site, os
path, norm = site.makepath('foo', 'bar')
print(path == os.path.join('foo', 'bar'))
print(norm == os.path.normcase(path))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_site_addpackage_reads_and_adds_relative_path() {
    let out = run_python(
        r#"
import site, sys, tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    sub_dir = os.path.join(tmpdir, 'rel_pkg')
    os.makedirs(sub_dir)
    pth_file = os.path.join(tmpdir, 'rel.pth')
    with open(pth_file, 'w') as f:
        f.write('rel_pkg\n')
    known_paths = set()
    site.addpackage(tmpdir, 'rel.pth', known_paths)
    print(sub_dir in sys.path)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_addpackage_ignores_comment_lines() {
    let out = run_python(
        r#"
import site, tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    pth_file = os.path.join(tmpdir, 'comment.pth')
    with open(pth_file, 'w') as f:
        f.write('# this is a comment\n')
    known_paths = set()
    site.addpackage(tmpdir, 'comment.pth', known_paths)
    print(tmpdir in known_paths)
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_site_addpackage_ignores_blank_lines() {
    let out = run_python(
        r#"
import site, tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    pth_file = os.path.join(tmpdir, 'blank.pth')
    with open(pth_file, 'w') as f:
        f.write('\n\n\n')
    known_paths = set()
    site.addpackage(tmpdir, 'blank.pth', known_paths)
    print(len(known_paths))
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_site_addsitedir_appends_to_end_of_sys_path() {
    let out = run_python(
        r#"
import site, sys, tempfile
with tempfile.TemporaryDirectory() as tmpdir:
    site.addsitedir(tmpdir)
    print(sys.path[-1] == tmpdir)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_site_main_function_prints_sys_path_info() {
    let out = run_python(
        r#"
import site, io, sys
buf = io.StringIO()
orig = sys.stdout
sys.stdout = buf
try:
    site.main()
finally:
    sys.stdout = orig
print('sys.path' in buf.getvalue())
"#,
    );
    assert_eq!(out, vec!["True"]);
}
