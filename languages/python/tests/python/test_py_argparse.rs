use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: argparse + optparse — argument parsing, subcommands, type conversion, defaults, required, nargs
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_argparse_basic_positional_and_optional() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser(description="Test")
parser.add_argument("name", help="Your name")
parser.add_argument("--age", type=int, default=0, help="Your age")
parser.add_argument("--verbose", action="store_true")

args = parser.parse_args(["Alice", "--age", "30", "--verbose"])
print(args.name)
print(args.age)
print(args.verbose)
"#;
    assert_eq!(run_python(src), vec!["Alice", "30", "True"]);
}

#[test]
fn test_py_argparse_default_values() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--host", default="localhost")
parser.add_argument("--port", type=int, default=8080)
parser.add_argument("--debug", action="store_false")

args = parser.parse_args([])
print(args.host)
print(args.port)
print(args.debug)
"#;
    assert_eq!(run_python(src), vec!["localhost", "8080", "True"]);
}

#[test]
fn test_py_argparse_nargs() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("files", nargs="+")
parser.add_argument("--tags", nargs="*", default=[])

args = parser.parse_args(["a.txt", "b.txt", "--tags", "web", "api"])
print(args.files)
print(args.tags)
"#;
    assert_eq!(
        run_python(src),
        vec!["['a.txt', 'b.txt']", "['web', 'api']"]
    );
}

#[test]
fn test_py_argparse_choices() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--level", choices=["debug", "info", "warning", "error"], default="info")
parser.add_argument("--format", choices=["json", "text", "csv"])

args = parser.parse_args(["--level", "debug", "--format", "json"])
print(args.level)
print(args.format)
"#;
    assert_eq!(run_python(src), vec!["debug", "json"]);
}

#[test]
fn test_py_argparse_type_conversion() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--count", type=int)
parser.add_argument("--ratio", type=float)

args = parser.parse_args(["--count", "42", "--ratio", "3.14"])
print(type(args.count).__name__)
print(type(args.ratio).__name__)
print(args.count * 2)
print(round(args.ratio, 2))
"#;
    assert_eq!(run_python(src), vec!["int", "float", "84", "3.14"]);
}

#[test]
fn test_py_argparse_subcommands() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
subparsers = parser.add_subparsers(dest="command")

push_parser = subparsers.add_parser("push")
push_parser.add_argument("remote", default="origin", nargs="?")

pull_parser = subparsers.add_parser("pull")
pull_parser.add_argument("--rebase", action="store_true")

args = parser.parse_args(["push", "upstream"])
print(args.command)
print(args.remote)

args2 = parser.parse_args(["pull", "--rebase"])
print(args2.command)
print(args2.rebase)
"#;
    assert_eq!(run_python(src), vec!["push", "upstream", "pull", "True"]);
}

#[test]
fn test_py_argparse_required_argument() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--output", required=True)

try:
    parser.parse_args([])
except SystemExit as e:
    print(f"SystemExit: {e.code}")
"#;
    assert_eq!(run_python(src), vec!["SystemExit: 2"]);
}

#[test]
fn test_py_argparse_append_action() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--item", action="append", dest="items")

args = parser.parse_args(["--item", "a", "--item", "b", "--item", "c"])
print(args.items)
"#;
    assert_eq!(run_python(src), vec!["['a', 'b', 'c']"]);
}

#[test]
fn test_py_argparse_mutually_exclusive_group() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group()
group.add_argument("--verbose", action="store_true")
group.add_argument("--quiet", action="store_true")

args = parser.parse_args(["--verbose"])
print(args.verbose)
print(args.quiet)

try:
    parser.parse_args(["--verbose", "--quiet"])
except SystemExit:
    print("Cannot use both")
"#;
    assert_eq!(run_python(src), vec!["True", "False", "Cannot use both"]);
}

#[test]
fn test_py_argparse_argument_groups() {
    let src = r#"
import argparse

parser = argparse.ArgumentParser()
network = parser.add_argument_group("network options")
network.add_argument("--host", default="localhost")
network.add_argument("--port", type=int, default=8080)

logging = parser.add_argument_group("logging options")
logging.add_argument("--log-level", default="INFO")

args = parser.parse_args(["--host", "example.com", "--port", "443"])
print(args.host)
print(args.port)
print(args.log_level)
"#;
    assert_eq!(run_python(src), vec!["example.com", "443", "INFO"]);
}
