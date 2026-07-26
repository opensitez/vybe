use super::helpers::run_python;

#[test]
fn test_python_format_width_and_fill() {
    let src = r#"
print('{:0>8}'.format(123))
print('{name:<8}|'.format(name='x'))
"#;
    assert_eq!(run_python(src), vec!["00000123", "x       |"]);
}

#[test]
fn test_python_format_nested_mapping() {
    let src = r#"
vals = {'x': 2, 'y': 3}
print('{x}+{y}={total}'.format(x=vals['x'], y=vals['y'], total=vals['x'] + vals['y']))
print('{0[0]} {0[1]}'.format(['a', 'b']))
"#;
    assert_eq!(run_python(src), vec!["2+3=5", "a b"]);
}
