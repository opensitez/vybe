// Python cmd module — Cmd REPL base, do_ methods, intro/prompt
use super::helpers::run_python;

#[test]
fn test_cmd_default_prompt() {
    let script = r#"
import cmd
c = cmd.Cmd()
print(c.prompt)
print(c.ruler)
"#;
    assert_eq!(run_python(script), vec!["(Cmd) ", "="]);
}

#[test]
fn test_cmd_custom_do_method() {
    let script = r#"
import cmd, io, sys

class Greeter(cmd.Cmd):
    prompt = "> "
    def do_hello(self, arg):
        print(f"Hello, {arg}!")
    def do_quit(self, arg):
        return True

g = Greeter(stdin=io.StringIO("hello world\nquit\n"), stdout=io.StringIO())
g.use_rawinput = False
g.cmdloop()
"#;
    // When output goes to StringIO the print inside do_hello still goes to stdout
    // So we test the behavior separately
    assert_eq!(run_python(script), vec!["Hello, world!"]);
}

#[test]
fn test_cmd_parseline() {
    let script = r#"
import cmd
c = cmd.Cmd()
cmd_name, args, line = c.parseline("hello world")
print(cmd_name)
print(args)
"#;
    assert_eq!(run_python(script), vec!["hello", "world"]);
}

#[test]
fn test_cmd_get_names() {
    let script = r#"
import cmd

class MyCli(cmd.Cmd):
    def do_foo(self, arg):
        pass
    def do_bar(self, arg):
        pass

names = MyCli.get_names()
print('do_foo' in names)
print('do_bar' in names)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_cmd_default_method() {
    let script = r#"
import cmd, io

output = io.StringIO()

class MyCli(cmd.Cmd):
    def default(self, line):
        print(f"unknown: {line}")
    def do_quit(self, arg):
        return True

cli = MyCli(stdin=io.StringIO("foobar\nquit\n"), stdout=io.StringIO())
cli.use_rawinput = False
cli.cmdloop()
"#;
    assert_eq!(run_python(script), vec!["unknown: foobar"]);
}
