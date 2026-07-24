use super::helpers::run_python;

// stat — S_IS* predicates, filemode, mode bit constants, os.stat integration

#[test]
fn test_stat_isreg_on_temp_file() {
    let out = run_python(r#"
import stat, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.close()
mode = os.stat(f.name).st_mode
print(stat.S_ISREG(mode))
os.unlink(f.name)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_stat_isdir_on_tempdir() {
    let out = run_python(r#"
import stat, tempfile, os, shutil
d = tempfile.mkdtemp()
mode = os.stat(d).st_mode
print(stat.S_ISDIR(mode))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_stat_filemode_regular_file() {
    let out = run_python(r#"
import stat, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.close()
mode = os.stat(f.name).st_mode
fm = stat.filemode(mode)
print(fm[0])   # first char '-' for regular file
os.unlink(f.name)
"#);
    assert_eq!(out, vec!["-"]);
}

#[test]
fn test_stat_filemode_directory() {
    let out = run_python(r#"
import stat, tempfile, os, shutil
d = tempfile.mkdtemp()
mode = os.stat(d).st_mode
fm = stat.filemode(mode)
print(fm[0])   # 'd' for directory
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["d"]);
}

#[test]
fn test_stat_imode_extracts_permission_bits() {
    let out = run_python(r#"
import stat
raw_mode = 0o100644
print(oct(stat.S_IMODE(raw_mode)))
"#);
    assert_eq!(out, vec!["0o644"]);
}

#[test]
fn test_stat_ifmt_extracts_type_bits() {
    let out = run_python(r#"
import stat
raw_mode = 0o100644
print(oct(stat.S_IFMT(raw_mode)))
"#);
    assert_eq!(out, vec!["0o100000"]);
}

#[test]
fn test_stat_constant_ifreg_value() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IFREG))
"#);
    assert_eq!(out, vec!["0o100000"]);
}

#[test]
fn test_stat_constant_ifdir_value() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IFDIR))
"#);
    assert_eq!(out, vec!["0o40000"]);
}

#[test]
fn test_stat_irusr_iwusr_ixusr_values() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IRUSR))
print(oct(stat.S_IWUSR))
print(oct(stat.S_IXUSR))
"#);
    assert_eq!(out, vec!["0o400", "0o200", "0o100"]);
}

#[test]
fn test_stat_irgrp_iwgrp_ixgrp_values() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IRGRP))
print(oct(stat.S_IWGRP))
print(oct(stat.S_IXGRP))
"#);
    assert_eq!(out, vec!["0o40", "0o20", "0o10"]);
}

#[test]
fn test_stat_iroth_iwoth_ixoth_values() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IROTH))
print(oct(stat.S_IWOTH))
print(oct(stat.S_IXOTH))
"#);
    assert_eq!(out, vec!["0o4", "0o2", "0o1"]);
}

#[test]
fn test_stat_irwxu_mask() {
    let out = run_python(r#"
import stat
print(oct(stat.S_IRWXU))
"#);
    assert_eq!(out, vec!["0o700"]);
}

#[test]
fn test_stat_st_size_index_constant() {
    let out = run_python(r#"
import stat
print(stat.ST_SIZE)
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_stat_st_mtime_index_constant() {
    let out = run_python(r#"
import stat
print(stat.ST_MTIME)
"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_stat_isreg_false_for_directory_mode() {
    let out = run_python(r#"
import stat
print(stat.S_ISREG(0o40755))
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_stat_isdir_false_for_regular_mode() {
    let out = run_python(r#"
import stat
print(stat.S_ISDIR(0o100644))
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_stat_filemode_string_length() {
    let out = run_python(r#"
import stat
fm = stat.filemode(0o100644)
print(len(fm))
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_stat_filemode_readable_bits() {
    let out = run_python(r#"
import stat
fm = stat.filemode(0o100644)
# owner read=1, write=1, execute=0; group read=1; other read=1
print(fm)
"#);
    assert_eq!(out, vec!["-rw-r--r--"]);
}

#[test]
fn test_stat_file_size_from_os_stat() {
    let out = run_python(r#"
import stat, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"hello world")
f.close()
info = os.stat(f.name)
print(info[stat.ST_SIZE])
os.unlink(f.name)
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_stat_isfifo_false_for_regular() {
    let out = run_python(r#"
import stat
print(stat.S_ISFIFO(0o100644))
"#);
    assert_eq!(out, vec!["False"]);
}
