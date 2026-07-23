use super::helpers::run_csharp;

#[test]
fn switch_expression_shape_and_sign_patterns() {
    let cases = [
        (0, "zero"),
        (1, "small-positive"),
        (9, "small-positive"),
        (10, "positive"),
        (-1, "negative"),
        (-9, "negative"),
    ];

    for (input, expected) in cases {
        let src = format!(
            r#"int value = {input}; string bucket = value switch {{ 0 => "zero", > 0 and < 10 => "small-positive", > 0 => "positive", _ => "negative" }}; Console.WriteLine(bucket);"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn switch_statement_with_type_patterns_and_when_guards() {
    let cases = [
        ("(object)0", "int-zero"),
        ("(object)42", "int-positive"),
        ("(object)\"A\".ToString()", "string-short"),
        ("(object)\"Long text value\"", "string-long"),
        ("(object)System.Math.PI", "double"),
        ("null", "null"),
    ];

    for (expression, expected) in cases {
        let src = format!(
            r#"object value = {expression};
string shape;
switch (value) {{
    case int value when value == 0: shape = "int-zero"; break;
    case int value: shape = "int-positive"; break;
    case string text when text.Length <= 5: shape = "string-short"; break;
    case string _: shape = "string-long"; break;
    case double: shape = "double"; break;
    case null: shape = "null"; break;
    default: shape = "other"; break;
}}
Console.WriteLine(shape);"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn tuple_patterns_split_points_by_quadrants() {
    let points = [
        (0, 0, "origin"),
        (0, 5, "axis"),
        (3, 0, "axis"),
        (2, 2, "first"),
        (2, -2, "fourth"),
        (-3, 4, "second"),
        (-1, -1, "third"),
    ];

    for (x, y, expected) in points {
        let src = format!(
            r#"
(int, int) point = ({x}, {y});
string quadrant = point switch {{
    (0, 0) => "origin",
    (0, _) => "axis",
    (_, 0) => "axis",
    (var a, var b) when a > 0 && b > 0 => "first",
    (var a, var b) when a > 0 && b < 0 => "fourth",
    (var a, var b) when a < 0 && b > 0 => "second",
    _ => "third"
}};
Console.WriteLine(quadrant);
"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn property_like_deconstruction_pattern_matrix() {
    let cases = [
        (1, 2, "same-parity"),
        (2, 4, "same-parity"),
        (2, 3, "mixed-parity"),
        (7, 8, "mixed-parity"),
        (0, 0, "same-parity"),
    ];

    for (left, right, _expected) in cases {
        let expected_bool = if (left % 2) == (right % 2) {
            "same-parity"
        } else {
            "mixed-parity"
        };
        let src = format!(
            r#"
(int Left, int Right) pair = ({left}, {right});
bool sameParity = (pair.Left & 1) == (pair.Right & 1);
Console.WriteLine(sameParity ? "same-parity" : "mixed-parity");
"#
        );
        assert_eq!(run_csharp(&src), vec![expected_bool.to_string()]);
    }
}
