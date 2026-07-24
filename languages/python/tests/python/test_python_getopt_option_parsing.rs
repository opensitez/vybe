use super::helpers::run_python;

// getopt — short options, long options, gnu_getopt, error handling

#[test]
fn test_getopt_short_flag_no_value() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-v"], "v")
print(opts)
print(args)
"#);
    assert_eq!(out, vec!["[('-v', '')]", "[]"]);
}

#[test]
fn test_getopt_short_option_with_value() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-o", "output.txt"], "o:")
print(opts)
"#);
    assert_eq!(out, vec!["[('-o', 'output.txt')]"]);
}

#[test]
fn test_getopt_multiple_short_options() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-v", "-n", "3"], "vn:")
print(opts)
"#);
    assert_eq!(out, vec!["[('-v', ''), ('-n', '3')]"]);
}

#[test]
fn test_getopt_long_option_flag() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["--verbose"], "", ["verbose"])
print(opts)
"#);
    assert_eq!(out, vec!["[('--verbose', '')]"]);
}

#[test]
fn test_getopt_long_option_with_equals_value() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["--output=file.txt"], "", ["output="])
print(opts)
"#);
    assert_eq!(out, vec!["[('--output', 'file.txt')]"]);
}

#[test]
fn test_getopt_long_option_with_separate_value() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["--output", "file.txt"], "", ["output="])
print(opts)
"#);
    assert_eq!(out, vec!["[('--output', 'file.txt')]"]);
}

#[test]
fn test_getopt_remaining_args_after_options() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-v", "file1", "file2"], "v")
print(args)
"#);
    assert_eq!(out, vec!["['file1', 'file2']"]);
}

#[test]
fn test_getopt_double_dash_stops_option_parsing() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-v", "--", "-x"], "v")
print(opts)
print(args)
"#);
    assert_eq!(out, vec!["[('-v', '')]", "['-x']"]);
}

#[test]
fn test_getopt_unknown_short_option_raises_error() {
    let out = run_python(r#"
import getopt
try:
    getopt.getopt(["-z"], "v")
except getopt.GetoptError as e:
    print(e.opt)
"#);
    assert_eq!(out, vec!["z"]);
}

#[test]
fn test_getopt_unknown_long_option_raises_error() {
    let out = run_python(r#"
import getopt
try:
    getopt.getopt(["--unknown"], "", ["verbose"])
except getopt.GetoptError as e:
    print("GetoptError")
"#);
    assert_eq!(out, vec!["GetoptError"]);
}

#[test]
fn test_getopt_missing_required_arg_raises_error() {
    let out = run_python(r#"
import getopt
try:
    getopt.getopt(["-o"], "o:")
except getopt.GetoptError as e:
    print("GetoptError")
"#);
    assert_eq!(out, vec!["GetoptError"]);
}

#[test]
fn test_getopt_gnu_getopt_mixed_args() {
    let out = run_python(r#"
import getopt
opts, args = getopt.gnu_getopt(["file.txt", "-v", "other"], "v")
print(opts)
print(sorted(args))
"#);
    assert_eq!(out, vec!["[('-v', '')]", "['file.txt', 'other']"]);
}

#[test]
fn test_getopt_gnu_getopt_long_mixed() {
    let out = run_python(r#"
import getopt
opts, args = getopt.gnu_getopt(["a.txt", "--verbose", "b.txt"], "", ["verbose"])
print(opts)
print(sorted(args))
"#);
    assert_eq!(out, vec!["[('--verbose', '')]", "['a.txt', 'b.txt']"]);
}

#[test]
fn test_getopt_combined_short_and_long() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-v", "--output=out.txt"], "v", ["output="])
print(opts)
"#);
    assert_eq!(out, vec!["[('-v', ''), ('--output', 'out.txt')]"]);
}

#[test]
fn test_getopt_error_msg_attribute() {
    let out = run_python(r#"
import getopt
try:
    getopt.getopt(["-z"], "")
except getopt.GetoptError as e:
    print(len(e.msg) > 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_getopt_empty_argv() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt([], "v")
print(opts)
print(args)
"#);
    assert_eq!(out, vec!["[]", "[]"]);
}

#[test]
fn test_getopt_concatenated_short_flags() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["-vn"], "vn")
print(opts)
"#);
    assert_eq!(out, vec!["[('-v', ''), ('-n', '')]"]);
}

#[test]
fn test_getopt_long_option_ambiguous_raises_error() {
    let out = run_python(r#"
import getopt
try:
    getopt.getopt(["--ver"], "", ["verbose", "version"])
except getopt.GetoptError:
    print("GetoptError")
"#);
    assert_eq!(out, vec!["GetoptError"]);
}

#[test]
fn test_getopt_only_long_opts_no_short() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["--dry-run"], "", ["dry-run"])
print(opts)
"#);
    assert_eq!(out, vec!["[('--dry-run', '')]"]);
}

#[test]
fn test_getopt_multiple_long_opts_same_call() {
    let out = run_python(r#"
import getopt
opts, args = getopt.getopt(["--foo", "--bar=baz"], "", ["foo", "bar="])
print(opts)
"#);
    assert_eq!(out, vec!["[('--foo', ''), ('--bar', 'baz')]"]);
}
