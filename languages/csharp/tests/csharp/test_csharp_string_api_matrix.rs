use super::helpers::run_csharp;

fn quote(value: &str) -> String {
    format!("{:?}", value)
}

fn bool_text(v: bool) -> &'static str {
    if v { "True" } else { "False" }
}

#[test]
fn string_matrix_basic_queries() {
    let words = [
        "", "a", "A", "ab", "abc", "foo", "bar", "baz", "xYz", "JSON", " space ", "Test",
    ];

    for left in words {
        for right in words {
            let len_src = format!(
                "string a = {}; string b = {}; Console.WriteLine((a + b).Length);",
                quote(left),
                quote(right)
            );
            assert_eq!(
                run_csharp(&len_src),
                &[(left.len() + right.len()).to_string().as_str()]
            );

            let contains_src = format!(
                "string a = {}; string b = {}; Console.WriteLine(a.Contains(b));",
                quote(left),
                quote(right)
            );
            assert_eq!(
                run_csharp(&contains_src),
                &[bool_text(left.contains(right))]
            );

            let starts_src = format!(
                "string a = {}; string b = {}; Console.WriteLine(a.StartsWith(b));",
                quote(left),
                quote(right)
            );
            assert_eq!(
                run_csharp(&starts_src),
                &[bool_text(left.starts_with(right))]
            );

            let ends_src = format!(
                "string a = {}; string b = {}; Console.WriteLine(a.EndsWith(b));",
                quote(left),
                quote(right)
            );
            assert_eq!(run_csharp(&ends_src), &[bool_text(left.ends_with(right))]);
        }
    }
}

#[test]
fn string_matrix_case_and_trim_behaviors() {
    let words = [
        "", "a", "Abc", "foo", "Bar", " xYz ", "  ", "upper", "Lower", "JSON", "test", "mix-ed",
    ];

    for word in words {
        let upper_src = format!(
            "string a = {}; Console.WriteLine(a.ToUpperInvariant());",
            quote(word)
        );
        assert_eq!(run_csharp(&upper_src), &[word.to_uppercase().as_str()]);

        let lower_src = format!(
            "string a = {}; Console.WriteLine(a.ToLowerInvariant());",
            quote(word)
        );
        assert_eq!(run_csharp(&lower_src), &[word.to_lowercase().as_str()]);

        let trim_src = format!(
            "string a = {}; Console.WriteLine('[' + a.Trim() + ']');",
            quote(word)
        );
        let expected_trim = format!("[{}]", word.trim());
        assert_eq!(run_csharp(&trim_src), &[expected_trim.as_str()]);

        let trim_start_src = format!(
            "string a = {}; Console.WriteLine('[' + a.TrimStart() + ']');",
            quote(word)
        );
        let expected_trim_start = format!("[{}]", word.trim_start());
        assert_eq!(run_csharp(&trim_start_src), &[expected_trim_start.as_str()]);

        let trim_end_src = format!(
            "string a = {}; Console.WriteLine('[' + a.TrimEnd() + ']');",
            quote(word)
        );
        let expected_trim_end = format!("[{}]", word.trim_end());
        assert_eq!(run_csharp(&trim_end_src), &[expected_trim_end.as_str()]);
    }
}

#[test]
fn string_matrix_replace_index_and_substring() {
    let words = [
        "", "a", "aa", "aba", "alpha", "beta", "gamma", "foo", "bar", "xyz", "xYz", "end",
    ];
    let needles = ["", "a", "b", "x", "al", "ta", " "];

    for source in words {
        for needle in needles {
            let replace_src = format!(
                r#"string s = {}; string p = {}; Console.WriteLine(s.Replace(p, "").Length);"#,
                quote(source),
                quote(needle)
            );
            let replaced = source.replace(needle, "");
            assert_eq!(
                run_csharp(&replace_src),
                &[replaced.len().to_string().as_str()]
            );
        }

        if source.len() > 0 {
            let first_src = format!(
                "string s = {}; Console.WriteLine(s.Substring(0, Math.Min(1, s.Length)));",
                quote(source)
            );
            let first = source.chars().next().unwrap_or_default().to_string();
            assert_eq!(run_csharp(&first_src), &[first.as_str()]);

            let last_src = format!(
                "string s = {}; Console.WriteLine(s.Substring(s.Length - 1, 1));",
                quote(source)
            );
            let last = source.chars().last().unwrap_or_default().to_string();
            assert_eq!(run_csharp(&last_src), &[last.as_str()]);

            let mid_src = format!(
                "string s = {}; Console.WriteLine(s.Substring(s.Length / 2, 1));",
                quote(source)
            );
            let mid_index = source.len() / 2;
            let mid = source[mid_index..mid_index + 1].to_string();
            assert_eq!(run_csharp(&mid_src), &[mid.as_str()]);
        }

        let split_src = format!(
            "string s = {}; string[] pieces = s.Split(' '); Console.WriteLine(pieces.Length);",
            quote(source)
        );
        let split_len = source.split(' ').count().to_string();
        assert_eq!(run_csharp(&split_src), &[split_len.as_str()]);
    }
}
