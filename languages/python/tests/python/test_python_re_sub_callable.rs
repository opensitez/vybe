use super::helpers::run_python;

#[test]
fn test_python_re_sub_with_callable() {
    let src = r#"
import re

def repl(m):
    return str(int(m.group(0)) * 2)

print(re.sub(r'\d+', repl, 'a1 b22'))
"#;
    assert_eq!(run_python(src), vec!["a2 b44"]);
}

#[test]
fn test_python_re_sub_count_and_flags() {
    let src = r#"
import re
print(re.sub(r'a', 'A', 'banana', count=2))
print(re.sub('x', 'y', 'XX', flags=re.IGNORECASE))
"#;
    assert_eq!(run_python(src), vec!["bAnAna", "YY"]);
}
