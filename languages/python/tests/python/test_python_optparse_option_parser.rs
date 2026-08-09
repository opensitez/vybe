use super::helpers::run_python;

// optparse — OptionParser, actions, types, defaults, groups

#[test]
fn test_optparse_store_string_option() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--name", dest="name", default="world")
opts, args = p.parse_args(["--name", "Alice"])
print(opts.name)
"#,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn test_optparse_default_value_used_when_absent() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--count", dest="count", type="int", default=5)
opts, _ = p.parse_args([])
print(opts.count)
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_optparse_store_true_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-v", "--verbose", action="store_true", dest="verbose", default=False)
opts, _ = p.parse_args(["-v"])
print(opts.verbose)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_optparse_store_false_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--no-debug", action="store_false", dest="debug", default=True)
opts, _ = p.parse_args(["--no-debug"])
print(opts.debug)
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_optparse_store_int_type() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-n", type="int", dest="num")
opts, _ = p.parse_args(["-n", "42"])
print(opts.num)
print(type(opts.num).__name__)
"#,
    );
    assert_eq!(out, vec!["42", "int"]);
}

#[test]
fn test_optparse_store_float_type() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--ratio", type="float", dest="ratio")
opts, _ = p.parse_args(["--ratio", "3.14"])
print(round(opts.ratio, 2))
"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn test_optparse_append_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-f", action="append", dest="files")
opts, _ = p.parse_args(["-f", "a.txt", "-f", "b.txt"])
print(opts.files)
"#,
    );
    assert_eq!(out, vec!["['a.txt', 'b.txt']"]);
}

#[test]
fn test_optparse_count_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-v", action="count", dest="verbosity", default=0)
opts, _ = p.parse_args(["-v", "-v", "-v"])
print(opts.verbosity)
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_optparse_choice_type_valid() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--fmt", type="choice", choices=["json", "xml", "csv"])
opts, _ = p.parse_args(["--fmt", "json"])
print(opts.fmt)
"#,
    );
    assert_eq!(out, vec!["json"]);
}

#[test]
fn test_optparse_choice_type_invalid_raises_error() {
    let out = run_python(
        r#"
import optparse, sys, io
p = optparse.OptionParser()
p.add_option("--fmt", type="choice", choices=["json", "xml"])
try:
    opts, _ = p.parse_args(["--fmt", "yaml"])
except SystemExit as e:
    print(e.code != 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_optparse_remaining_positional_args() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-v", action="store_true", dest="verbose")
opts, args = p.parse_args(["-v", "file1", "file2"])
print(args)
"#,
    );
    assert_eq!(out, vec!["['file1', 'file2']"]);
}

#[test]
fn test_optparse_dest_attribute_name() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--input-file", dest="input_file")
opts, _ = p.parse_args(["--input-file", "data.txt"])
print(opts.input_file)
"#,
    );
    assert_eq!(out, vec!["data.txt"]);
}

#[test]
fn test_optparse_option_group_title() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
g = optparse.OptionGroup(p, "Advanced Options")
g.add_option("--debug", action="store_true", dest="debug")
p.add_option_group(g)
opts, _ = p.parse_args(["--debug"])
print(opts.debug)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_optparse_short_and_long_equivalent() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-o", "--output", dest="output")
o1, _ = p.parse_args(["-o", "a.txt"])
o2, _ = p.parse_args(["--output", "a.txt"])
print(o1.output == o2.output)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_optparse_store_const_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--mode", action="store_const", const="fast", dest="mode", default="slow")
opts, _ = p.parse_args(["--mode"])
print(opts.mode)
"#,
    );
    assert_eq!(out, vec!["fast"]);
}

#[test]
fn test_optparse_int_type_coercion_error_exits() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("-n", type="int", dest="n")
try:
    p.parse_args(["-n", "notanint"])
except SystemExit as e:
    print(e.code != 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_optparse_multiple_option_groups() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
g1 = optparse.OptionGroup(p, "Group 1")
g1.add_option("--foo", action="store_true", dest="foo")
g2 = optparse.OptionGroup(p, "Group 2")
g2.add_option("--bar", action="store_true", dest="bar")
p.add_option_group(g1)
p.add_option_group(g2)
opts, _ = p.parse_args(["--foo", "--bar"])
print(opts.foo, opts.bar)
"#,
    );
    assert_eq!(out, vec!["True True"]);
}

#[test]
fn test_optparse_none_default_for_unset_option() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--name", dest="name")
opts, _ = p.parse_args([])
print(opts.name)
"#,
    );
    assert_eq!(out, vec!["None"]);
}

#[test]
fn test_optparse_append_const_action() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
p.add_option("--json", action="append_const", const="json", dest="formats")
p.add_option("--xml",  action="append_const", const="xml",  dest="formats")
opts, _ = p.parse_args(["--json", "--xml"])
print(sorted(opts.formats))
"#,
    );
    assert_eq!(out, vec!["['json', 'xml']"]);
}

#[test]
fn test_optparse_error_on_unknown_option() {
    let out = run_python(
        r#"
import optparse
p = optparse.OptionParser()
try:
    p.parse_args(["--unknown-flag"])
except SystemExit as e:
    print(e.code != 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
