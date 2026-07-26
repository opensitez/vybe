use super::helpers::run_python;

#[test]
fn test_python_enum_flag_operations() {
    let src = r#"
from enum import Flag, auto

class Opt(Flag):
    READ = auto()
    WRITE = auto()
    EXEC = auto()

perm = Opt.READ | Opt.EXEC
print(perm)
print(perm.value == (Opt.READ.value | Opt.EXEC.value))
"#;
    assert_eq!(run_python(src), vec!["Opt.READ|EXEC", "True"]);
}

#[test]
fn test_python_enum_member_name_lookup() {
    let src = r#"
from enum import IntFlag, auto

class Mode(IntFlag):
    A = auto()
    B = auto()

m = Mode.A | Mode.B
print(int(m))
print(m & Mode.B == Mode.B)
"#;
    assert_eq!(run_python(src), vec!["3", "True"]);
}
