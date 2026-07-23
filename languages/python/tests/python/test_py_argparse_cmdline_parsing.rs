use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Argparse Command Line Parsing — positionals, optionals, defaults, type conversion, subcommands, nargs
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_argparse_positional_and_flags() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("filename")
parser.add_argument("-v", "--verbose", action="store_true")
parser.add_argument("-n", "--count", type=int, default=1)

args = parser.parse_args(["data.csv", "-v", "-n", "5"])
print(args.filename)
print(args.verbose)
print(args.count)
"#;
    assert_eq!(run_python(src), vec!["data.csv", "True", "5"]);
}

#[test]
fn test_py_argparse_nargs_list_accumulation() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--items", nargs="+", type=int)
parser.add_argument("--tags", nargs="*", default=[])

args = parser.parse_args(["--items", "10", "20", "30", "--tags", "a", "b"])
print(args.items)
print(args.tags)
"#;
    assert_eq!(run_python(src), vec!["[10, 20, 30]", "['a', 'b']"]);
}

#[test]
fn test_py_argparse_choices_validation() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--mode", choices=["fast", "slow", "auto"], default="auto")

args = parser.parse_args(["--mode", "fast"])
print(args.mode)
"#;
    assert_eq!(run_python(src), vec!["fast"]);
}

#[test]
fn test_py_argparse_subparsers_command_routing() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
subparsers = parser.add_subparsers(dest="subcommand")

build_cmd = subparsers.add_parser("build")
build_cmd.add_argument("--target", default="all")

test_cmd = subparsers.add_parser("test")
test_cmd.add_argument("--filter", default="*")

args = parser.parse_args(["build", "--target", "web"])
print(args.subcommand)
print(args.target)
"#;
    assert_eq!(run_python(src), vec!["build", "web"]);
}

#[test]
fn test_py_argparse_action_append_and_count() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-v", action="count", default=0)
parser.add_argument("-i", "--include", action="append")

args = parser.parse_args(["-vvv", "-i", "dir1", "-i", "dir2"])
print(args.v)
print(args.include)
"#;
    assert_eq!(run_python(src), vec!["3", "['dir1', 'dir2']"]);
}

#[test]
fn test_py_argparse_mutually_exclusive_group() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument("--json", action="store_true")
group.add_argument("--yaml", action="store_true")

args = parser.parse_args(["--json"])
print(args.json)
print(args.yaml)
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_argparse_custom_type_converter() {
    let src = r#"
import argparse

def parse_pair(s):
    k, v = s.split("=")
    return k, int(v)

parser = argparse.ArgumentParser()
parser.add_argument("--set", type=parse_pair)

args = parser.parse_args(["--set", "limit=100"])
print(args.set)
"#;
    assert_eq!(run_python(src), vec!["('limit', 100)"]);
}

#[test]
fn test_py_argparse_defaults_from_dict() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--host")
parser.add_argument("--port", type=int)
parser.set_defaults(host="localhost", port=8080)

args = parser.parse_args([])
print(args.host, args.port)
"#;
    assert_eq!(run_python(src), vec!["localhost 8080"]);
}

#[test]
fn test_py_argparse_dest_attribute_renaming() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--output-dir", dest="outdir")

args = parser.parse_args(["-o", "/tmp/build"])
print(args.outdir)
"#;
    assert_eq!(run_python(src), vec!["/tmp/build"]);
}

#[test]
fn test_py_argparse_store_const_action() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--env", action="store_const", const="production", default="development")

args = parser.parse_args(["--env"])
print(args.env)
"#;
    assert_eq!(run_python(src), vec!["production"]);
}
