use super::helpers::run_vb;

// Len, Trim, LTrim, RTrim
#[test]
fn string_len() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Len("Hello")): End Sub: End Module"#),
        vec!["5"]
    );
}
#[test]
fn string_trim() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Trim("  Hello  ") & "!"): End Sub: End Module"#
        ),
        vec!["Hello!"]
    );
}
#[test]
fn string_ltrim() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(LTrim("  Hello  ") & "!"): End Sub: End Module"#
        ),
        vec!["Hello  !"]
    );
}
#[test]
fn string_rtrim() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(RTrim("  Hello  ") & "!"): End Sub: End Module"#
        ),
        vec!["  Hello!"]
    );
}

// Mid, Left, Right
#[test]
fn string_mid_3() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Mid("Hello", 2, 3)): End Sub: End Module"#
        ),
        vec!["ell"]
    );
}
#[test]
fn string_mid_2() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Mid("Hello", 3)): End Sub: End Module"#),
        vec!["llo"]
    );
}
#[test]
fn string_left() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Left("Hello", 2)): End Sub: End Module"#),
        vec!["He"]
    );
}
#[test]
fn string_right() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Right("Hello", 3)): End Sub: End Module"#
        ),
        vec!["llo"]
    );
}

// InStr, InStrRev
#[test]
fn string_instr_start() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(InStr(1, "Hello World", "o")): End Sub: End Module"#
        ),
        vec!["5"]
    );
}
#[test]
fn string_instr_default() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(InStr("Hello World", "o")): End Sub: End Module"#
        ),
        vec!["5"]
    );
}
#[test]
fn string_instrrev_start() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(InStrRev("Hello World", "o", 11)): End Sub: End Module"#
        ),
        vec!["8"]
    );
}
#[test]
fn string_instrrev_default() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(InStrRev("Hello World", "o")): End Sub: End Module"#
        ),
        vec!["8"]
    );
}

// UCase, LCase
#[test]
fn string_ucase() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(UCase("Hello")): End Sub: End Module"#),
        vec!["HELLO"]
    );
}
#[test]
fn string_lcase() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(LCase("Hello")): End Sub: End Module"#),
        vec!["hello"]
    );
}

// StrComp
#[test]
fn string_strcomp_binary() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(StrComp("Hello", "hello", CompareMethod.Binary)): End Sub: End Module"#
        ),
        vec!["-1"]
    );
}
#[test]
fn string_strcomp_text() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(StrComp("Hello", "hello", CompareMethod.Text)): End Sub: End Module"#
        ),
        vec!["0"]
    );
}

// StrReverse
#[test]
fn string_strreverse() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(StrReverse("Hello")): End Sub: End Module"#
        ),
        vec!["olleH"]
    );
}

// Space, String
#[test]
fn string_space() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine("A" & Space(3) & "B"): End Sub: End Module"#
        ),
        vec!["A   B"]
    );
}
#[test]
fn string_string_char() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(StrDup(3, "X"c)): End Sub: End Module"#),
        vec!["XXX"]
    );
}

// Format
#[test]
fn string_format_currency() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Format(12.34, "Currency")): End Sub: End Module"#
        ),
        vec!["$12.34"]
    );
}
#[test]
fn string_format_percent() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Format(0.123, "Percent")): End Sub: End Module"#
        ),
        vec!["12.30%"]
    );
}
#[test]
fn string_format_custom() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Format(123.456, "0.00")): End Sub: End Module"#
        ),
        vec!["123.46"]
    );
}

// Asc, Chr, AscW, ChrW
#[test]
fn string_asc() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Asc("A")): End Sub: End Module"#),
        vec!["65"]
    );
}
#[test]
fn string_chr() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Chr(65)): End Sub: End Module"#),
        vec!["A"]
    );
}
#[test]
fn string_ascw() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(AscW("A")): End Sub: End Module"#),
        vec!["65"]
    );
}
#[test]
fn string_chrw() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(ChrW(65)): End Sub: End Module"#),
        vec!["A"]
    );
}

// Filter, Join, Split
#[test]
fn string_split() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Split("A,B,C", ",")(1)): End Sub: End Module"#
        ),
        vec!["B"]
    );
}
#[test]
fn string_join() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Join({"A", "B"}, "-")): End Sub: End Module"#
        ),
        vec!["A-B"]
    );
}
#[test]
fn string_filter_include() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Filter({"Apple", "Banana", "Apricot"}, "A").Length): End Sub: End Module"#
        ),
        vec!["2"]
    );
}
#[test]
fn string_filter_exclude() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Filter({"Apple", "Banana", "Apricot"}, "A", False).Length): End Sub: End Module"#
        ),
        vec!["1"]
    );
}

// Like Operator edge cases
#[test]
fn like_operator_star() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine("Hello" Like "H*"): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn like_operator_question() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine("Hello" Like "H?llo"): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn like_operator_number() {
    assert_eq!(
        run_vb(
            "Module M: Sub Main(): Console.WriteLine(\"123\" Like \"###\"): End Sub: End Module"
        ),
        vec!["True"]
    );
}
#[test]
fn like_operator_charlist() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine("C" Like "[A-Z]"): End Sub: End Module"#),
        vec!["True"]
    );
}
#[test]
fn like_operator_negcharlist() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine("1" Like "[!A-Z]"): End Sub: End Module"#
        ),
        vec!["True"]
    );
}

// String assignment Mid (Statement)
#[test]
fn string_mid_stmt() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim s = "Hello": Mid(s, 2, 2) = "XX": Console.WriteLine(s): End Sub: End Module"#
        ),
        vec!["HXXlo"]
    );
}
#[test]
fn string_mid_stmt_shorter() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim s = "Hello": Mid(s, 2, 2) = "X": Console.WriteLine(s): End Sub: End Module"#
        ),
        vec!["HXllo"]
    );
}

// IsNumeric, IsDate, IsDBNull, IsNothing, IsError, IsArray, IsReference
#[test]
fn type_isnumeric() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsNumeric("123.45")): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn type_isdate() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsDate("2020-01-01")): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn type_isdbnull() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsDBNull(System.DBNull.Value)): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn type_isnothing() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsNothing(Nothing)): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn type_iserror() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsError(New System.Exception())): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn type_isarray() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(IsArray({1})): End Sub: End Module"#),
        vec!["True"]
    );
}
#[test]
fn type_isreference() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(IsReference("String")): End Sub: End Module"#
        ),
        vec!["True"]
    );
}

// Replace
#[test]
fn string_replace() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Replace("Hello", "l", "w")): End Sub: End Module"#
        ),
        vec!["Hewwo"]
    );
}

// Str, Val
#[test]
fn string_str() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Str(42)): End Sub: End Module"#),
        vec![" 42"]
    );
}
#[test]
fn string_str_neg() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Str(-42)): End Sub: End Module"#),
        vec!["-42"]
    );
}
#[test]
fn string_val() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Val("  42  ")): End Sub: End Module"#),
        vec!["42"]
    );
}
#[test]
fn string_val_hex() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Val("&H10")): End Sub: End Module"#),
        vec!["16"]
    );
}
#[test]
fn string_val_oct() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(Val("&O10")): End Sub: End Module"#),
        vec!["8"]
    );
}
