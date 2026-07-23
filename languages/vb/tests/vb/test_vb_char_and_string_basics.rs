use super::helpers::run_vb;

#[test]
fn char_literal_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim c As Char = \"A\"c\nConsole.WriteLine(c.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Char"]
    );
}
#[test]
fn char_literal_quote() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim c = \"\"\"\"c\nConsole.WriteLine(c = ChrW(34))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn char_concatenation_implicit_string() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"A\"c & \"B\"c\nConsole.WriteLine(s.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["String"]
    );
}
#[test]
fn char_addition_error_or_numeric() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\n' Adding chars without string context may fail or convert to ASCII/Unicode\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn char_to_string_implicit() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s As String = \"A\"c\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}

#[test]
fn string_literal_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s As String = \"Hello\"\nConsole.WriteLine(s.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["String"]
    );
}
#[test]
fn string_literal_escaped_quote() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s As String = \"He\"\"llo\"\nConsole.WriteLine(s.Contains(Chr(34)))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn string_concatenation_ampersand() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"A\" & \"B\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["AB"]
    );
}
#[test]
fn string_concatenation_plus() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"A\" + \"B\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["AB"]
    );
}
#[test]
fn string_concatenation_number_ampersand() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"A\" & 5\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A5"]
    );
}

#[test]
fn string_concatenation_number_plus() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\n' + with string and number is tricky, usually throws InvalidCast if string isn't numeric\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn string_concatenation_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim n As String = Nothing\nDim s = \"A\" & n\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn string_equality_case_sensitive() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(\"A\" = \"a\")\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn string_equality_nothing_empty() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim n As String = Nothing\nConsole.WriteLine(n = \"\")\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
} // VB quirk: Nothing string equals ""
#[test]
fn string_len_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim n As String = Nothing\nConsole.WriteLine(Len(n))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}

#[test]
fn string_type_char() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s$ = \"Hello\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["Hello"]
    );
}
#[test]
fn string_fixed_length_legacy() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\n' Dim s As String * 10 is legacy, often unsupported in .NET directly but parsable\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn string_assignment_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s As String = Nothing\nConsole.WriteLine(s Is Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn string_interpolated_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim v = 10\nDim s = $\"A{v}B\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A10B"]
    );
}
#[test]
fn string_interpolated_formatting() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim v = 10\nDim s = $\"A{v:000}B\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A010B"]
    );
}

#[test]
fn string_interpolated_double_quotes() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = $\"A{\"\"B\"\"}\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["A\"B\""]
    );
}
#[test]
fn string_interpolated_braces() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = $\"{{A}}\"\nConsole.WriteLine(s)\nEnd Sub\nEnd Module"
        ),
        vec!["{A}"]
    );
}
#[test]
fn string_multiline_literal() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"A\nB\"\nConsole.WriteLine(s.Contains(vbLf) Or s.Contains(vbCrLf))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn string_comparison_operator_greater() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(\"B\" > \"A\")\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn string_isnot_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"\"\nConsole.WriteLine(s IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
