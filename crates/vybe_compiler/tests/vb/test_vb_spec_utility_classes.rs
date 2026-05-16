use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $body:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let output = run_vb($body);
            assert_eq!(output, vec![$expected.to_string()]);
        }
    };
}

#[test]
fn utility_spec_stringbuilder_mutations() {
    let output = run_vb(
        r#"
Imports System.Text
Module Program
    Sub Main()
        Dim sb As New StringBuilder("vy")
        sb.Append("be")
        sb.Insert(2, "-")
        sb.Replace("-", "")
        Console.WriteLine(sb.ToString())
        Console.WriteLine(sb.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(output, vec!["vybe", "4"]);
}

#[test]
fn utility_spec_stringbuilder_clear_remove_and_appendline() {
    let output = run_vb(
        r#"
Imports System.Text
Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendLine("alpha")
        sb.Append("beta")
        sb.Remove(0, 6)
        Console.WriteLine(sb.ToString())
        sb.Clear()
        Console.WriteLine(sb.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(output, vec!["beta", "0"]);
}

vb_expr_spec!(
    utility_spec_regex_is_match,
    r#"
Imports System.Text.RegularExpressions
Module Program
    Sub Main()
        Dim re As New Regex("a+")
        Console.WriteLine(re.IsMatch("caa"))
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    utility_spec_regex_replace,
    r#"
Imports System.Text.RegularExpressions
Module Program
    Sub Main()
        Dim re As New Regex("a+")
        Console.WriteLine(re.Replace("baac", "X"))
    End Sub
End Module
"#,
    "bXc"
);

vb_expr_spec!(
    utility_spec_regex_split,
    r#"
Imports System.Text.RegularExpressions
Module Program
    Sub Main()
        Dim re As New Regex(",")
        Dim parts() As String = re.Split("a,b,c")
        Console.WriteLine(Join(parts, "|"))
    End Sub
End Module
"#,
    "a|b|c"
);

vb_expr_spec!(
    utility_spec_regex_match_value,
    r#"
Imports System.Text.RegularExpressions
Module Program
    Sub Main()
        Dim re As New Regex("\d+")
        Console.WriteLine(re.Match("x12y").Value)
    End Sub
End Module
"#,
    "12"
);

vb_expr_spec!(
    utility_spec_regex_matches_count,
    r#"
Imports System.Text.RegularExpressions
Module Program
    Sub Main()
        Dim re As New Regex("\d+")
        Console.WriteLine(re.Matches("1 22 333").Count)
    End Sub
End Module
"#,
    "3"
);

vb_expr_spec!(
    utility_spec_convert_to_base64,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Convert.ToBase64String("hello"))
    End Sub
End Module
"#,
    "aGVsbG8="
);

vb_expr_spec!(
    utility_spec_convert_from_base64,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Convert.FromBase64String("aGVsbG8="))
    End Sub
End Module
"#,
    "hello"
);

vb_expr_spec!(
    utility_spec_convert_to_datetime,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Year(Convert.ToDateTime("2024-05-14")))
    End Sub
End Module
"#,
    "2024"
);

#[test]
fn utility_spec_stopwatch_state_transitions() {
    let output = run_vb(
        r#"
Imports System.Diagnostics
Module Program
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        Console.WriteLine(sw.IsRunning)
        sw.Stop()
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
        sw.Reset()
        Console.WriteLine(sw.ElapsedMilliseconds)
        sw.Restart()
        Console.WriteLine(sw.IsRunning)
    End Sub
End Module
"#,
    );

    assert_eq!(output, vec!["True", "False", "True", "0", "True"]);
}

vb_expr_spec!(
    utility_spec_stopwatch_startnew,
    r#"
Imports System.Diagnostics
Module Program
    Sub Main()
        Dim sw = Stopwatch.StartNew()
        Console.WriteLine(sw.IsRunning)
    End Sub
End Module
"#,
    "True"
);
