use super::helpers::run_python;

#[test]
fn test_python_tempfile_named_file() {
    let src = r#"
import tempfile, os
with tempfile.NamedTemporaryFile('w+', delete=False) as f:
    f.write('hi')
    path = f.name
with open(path, 'r') as f:
    print(f.read())
os.remove(path)
print('removed')
"#;
    assert_eq!(run_python(src), vec!["hi", "removed"]);
}

#[test]
fn test_python_tempfile_temporary_directory() {
    let src = r#"
import tempfile, os
with tempfile.TemporaryDirectory() as d:
    name = d + '/x.txt'
    with open(name, 'w') as f:
        f.write('ok')
    print(os.path.exists(name))
print('closed')
"#;
    assert_eq!(run_python(src), vec!["True", "closed"]);
}
