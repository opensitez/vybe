use super::helpers::run_csharp;

fn quote(value: Option<i32>) -> String {
    match value {
        Some(v) => v.to_string(),
        None => "null".to_string(),
    }
}

#[test]
fn nullable_int_coalesce_fallback_matrix() {
    let values: [Option<i32>; 8] = [Some(-3), Some(0), Some(17), Some(42), None, Some(5), None, Some(-1)];

    for value in values {
        let src = format!(
            "int? value = {}; Console.WriteLine(value ?? -1);",
            quote(value)
        );
        let expected = value.unwrap_or(-1).to_string();
        assert_eq!(run_csharp(&src), vec![expected]);
    }
}

#[test]
fn nullable_string_length_with_null_conditional_chain() {
    let cases = [
        "null",
        "\"\"",
        "\"a\"",
        "\"hello\"",
        "\"  trim  \"",
    ];

    for raw in cases {
        let src = format!(
            "string? value = {raw}; int? len = value?.Trim().Length; Console.WriteLine(len ?? -1);"
        );
        let expected = if raw == "null" {
            "-1".to_string()
        } else {
            let decoded = raw.trim_matches('"');
            decoded.trim().chars().count().to_string()
        };
        assert_eq!(run_csharp(&src), vec![expected]);
    }
}

#[test]
fn null_coalescing_assignment_on_nullable_int() {
    let pairs = [(None, 100), (Some(0), 9), (Some(7), 3), (None, -8)];

    for (initial, fallback) in pairs {
        let src = format!(
            r#"
int? value = {initial_value};
value ??= {fallback};
Console.WriteLine(value.Value);
"#,
            initial_value = match initial {
                Some(v) => v.to_string(),
                None => "null".to_string(),
            }
        );
        assert_eq!(run_csharp(&src), vec![initial.unwrap_or(fallback).to_string()]);
    }
}

#[test]
fn nullable_collection_count_and_sum_matrix() {
    let cases: Vec<Vec<Option<i32>>> = vec![
        vec![Some(1), Some(2), None, Some(4)],
        vec![None, None],
        vec![Some(-1), Some(-2), Some(-3)],
        vec![Some(10)],
        vec![],
    ];

    for values in cases {
        let literal = values
            .iter()
            .map(|value| match value {
                Some(v) => v.to_string(),
                None => "null".to_string(),
            })
            .collect::<Vec<_>>()
            .join(", ");

        let expected_count: usize = values.iter().filter(|value| value.is_some()).count();
        let expected_sum: i32 = values.iter().filter_map(|value| *value).sum();

        let src = format!(
            "int?[] values = new int?[] {{ {literal} }}; int count = 0; int sum = 0; foreach (var value in values) {{ if (value.HasValue) {{ count += 1; sum += value.Value; }} Console.WriteLine(count); Console.WriteLine(sum); }}"
        );
        assert_eq!(
            run_csharp(&src),
            vec![expected_count.to_string(), expected_sum.to_string()]
        );
    }
}

#[test]
fn nullable_value_type_member_and_hasvalue_paths() {
    let pairs = [
        (Some(1), "has"),
        (Some(-9), "has"),
        (None, "missing"),
    ];

    for (value, expected) in pairs {
        let src = format!(
            r#"int? value = {}; bool has = value.HasValue; Console.WriteLine(has ? "has" : "missing"); Console.WriteLine(value.GetValueOrDefault(123));"#,
            value.map_or("null".to_string(), |v| v.to_string())
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                expected.to_string(),
                value.unwrap_or(123).to_string()
            ]
        );
    }
}
