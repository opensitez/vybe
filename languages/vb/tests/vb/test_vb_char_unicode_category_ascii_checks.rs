use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Char Unicode Categories, ASCII & Utility Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_char_is_digit_is_letter() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsDigit("5"c) & "|" & Char.IsLetter("A"c) & "|" & Char.IsDigit("A"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_char_is_letter_or_digit() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsLetterOrDigit("9"c) & "|" & Char.IsLetterOrDigit("Z"c) & "|" & Char.IsLetterOrDigit("@"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_char_is_lower_is_upper() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsLower("a"c) & "|" & Char.IsUpper("A"c) & "|" & Char.IsUpper("a"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_char_is_white_space_and_control() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsWhiteSpace(" "c) & "|" & Char.IsWhiteSpace(vbTab) & "|" & Char.IsControl(vbLf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_char_is_punctuation_and_symbol() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsPunctuation("."c) & "|" & Char.IsPunctuation(","c) & "|" & Char.IsSymbol("+"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_char_is_separator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsSeparator(" "c) & "|" & Char.IsSeparator("A"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_char_get_numeric_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dVal = Char.GetNumericValue("7"c)
        Dim letterVal = Char.GetNumericValue("X"c)
        Console.WriteLine(dVal & "|" & letterVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["7|-1"]);
}

#[test]
fn test_vb_char_get_unicode_category() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim catUpper = Char.GetUnicodeCategory("A"c)
        Dim catDigit = Char.GetUnicodeCategory("1"c)
        Console.WriteLine(catUpper.ToString() & "|" & catDigit.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UppercaseLetter|DecimalDigitNumber"]);
}

#[test]
fn test_vb_char_is_ascii_checks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim cAscii = "A"c
        Console.WriteLine(Char.IsAscii(cAscii) & "|" & Char.IsAsciiDigit("5"c) & "|" & Char.IsAsciiLetter("Z"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_char_is_ascii_hex_digit() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Char.IsAsciiHexDigit("F"c) & "|" & Char.IsAsciiHexDigit("a"c) & "|" & Char.IsAsciiHexDigit("G"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_char_to_lower_to_upper_culture() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim lower = Char.ToLower("K"c, CultureInfo.InvariantCulture)
        Dim upper = Char.ToUpper("m"c, CultureInfo.InvariantCulture)
        Console.WriteLine(lower & "|" & upper)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["k|M"]);
}

#[test]
fn test_vb_char_is_surrogate_surrogate_pair() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim highSurrogate = ChrW(&HD83D)
        Dim lowSurrogate = ChrW(&HDE00)
        Console.WriteLine(Char.IsSurrogate(highSurrogate) & "|" & Char.IsSurrogatePair(highSurrogate, lowSurrogate))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_char_convert_to_utf32_surrogate_pair() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim highSurrogate = ChrW(&HD83D)
        Dim lowSurrogate = ChrW(&HDE00)
        Dim utf32 = Char.ConvertToUtf32(highSurrogate, lowSurrogate)
        Console.WriteLine(Hex(utf32))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1F600"]);
}

#[test]
fn test_vb_char_convert_from_utf32() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim str = Char.ConvertFromUtf32(&H1F600)
        Console.WriteLine(str.Length & "|" & Char.IsSurrogatePair(str, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|True"]);
}

#[test]
fn test_vb_char_is_high_low_surrogate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim high = ChrW(&HD83D)
        Dim low = ChrW(&HDE00)
        Console.WriteLine(Char.IsHighSurrogate(high) & "|" & Char.IsLowSurrogate(low))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_char_string_indexed_character_checks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim text = "A1 #"
        Console.WriteLine(Char.IsLetter(text, 0) & "|" & Char.IsDigit(text, 1) & "|" & Char.IsWhiteSpace(text, 2) & "|" & Char.IsPunctuation(text, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True|True"]);
}

#[test]
fn test_vb_char_comparison_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim c1 As Char = "A"c
        Dim c2 As Char = "B"c
        Console.WriteLine((c1 < c2) & "|" & (c1 = "A"c) & "|" & (c1 <> c2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_char_chrw_ascw_conversion() {
    let src = r#"
Module Program
    Sub Main()
        Dim code = AscW("Z"c)
        Dim ch = ChrW(code)
        Console.WriteLine(code & "|" & ch)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90|Z"]);
}

#[test]
fn test_vb_char_min_max_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(AscW(Char.MinValue) & "|" & AscW(Char.MaxValue))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|-1"]);
}

#[test]
fn test_vb_char_array_to_string_and_reverse() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim chars As Char() = "VisualBasic".ToCharArray()
        Array.Reverse(chars)
        Console.WriteLine(New String(chars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["cisaBlausiV"]);
}
