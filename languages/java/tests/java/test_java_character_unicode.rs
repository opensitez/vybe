use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(character_literal, "System.out.println('z');", "z");
jt!(character_codepoint, "System.out.println((int)'A');", "65");
jt!(
    character_to_uppercase,
    "char c = 'a'; System.out.println(Character.toUpperCase(c));",
    "A"
);
jt!(
    character_to_lowercase,
    "char c = 'Z'; System.out.println(Character.toLowerCase(c));",
    "z"
);
jt!(
    character_is_digit_true,
    "System.out.println(Character.isDigit('7'));",
    "true"
);
jt!(
    character_is_digit_false,
    "System.out.println(Character.isDigit('a'));",
    "false"
);
jt!(
    character_is_letter_true,
    "System.out.println(Character.isLetter('x'));",
    "true"
);
jt!(
    character_is_letter_false,
    "System.out.println(Character.isLetter('7'));",
    "false"
);
jt!(
    character_is_whitespace,
    "System.out.println(Character.isWhitespace(' '));",
    "true"
);
jt!(
    character_is_upper_case,
    "System.out.println(Character.isUpperCase('Q'));",
    "true"
);
jt!(
    character_is_lower_case,
    "System.out.println(Character.isLowerCase('q'));",
    "true"
);
jt!(
    character_is_letter_or_digit_true,
    "System.out.println(Character.isLetterOrDigit('9'));",
    "true"
);
jt!(
    character_is_letter_or_digit_false,
    "System.out.println(Character.isLetterOrDigit('!'));",
    "false"
);
jt!(
    character_is_high_surrogate,
    "System.out.println(Character.isHighSurrogate((char)0xD800));",
    "true"
);
jt!(
    character_is_low_surrogate,
    "System.out.println(Character.isLowSurrogate((char)0xDC00));",
    "true"
);
jt!(
    character_to_chars_len,
    "char[] out = Character.toChars(0x41); System.out.println(out.length);",
    "1"
);
jt!(
    character_to_chars_value,
    "char[] out = Character.toChars(0x41); System.out.println(out[0]);",
    "A"
);
jt!(
    unicode_escape_sequence_a,
    "System.out.println(\"\u{0041}\");",
    "A"
);
jt!(
    unicode_escape_sequence_b,
    "System.out.println(\"A\" + \"\u{0042}\");",
    "AB"
);
jt!(
    unicode_escape_numeric,
    "System.out.println(\"\u{0031}\" + \"2\");",
    "12"
);
jt!(
    unicode_escape_control,
    "System.out.println(\"\u{000A}\".equals(\"\\n\"));",
    "true"
);
jt!(
    unicode_surrogate_pair_supported,
    "System.out.println(Character.isSurrogatePair((char)0xD83D, (char)0xDE00));",
    "true"
);
jt!(
    unicode_to_string_from_cp,
    "String s = new String(Character.toChars(0x1F600)); System.out.println(s.codePointAt(0) == 0x1F600);",
    "true"
);
jt!(
    unicode_is_valid_code_point_true,
    "System.out.println(Character.isValidCodePoint(65));",
    "true"
);
jt!(
    unicode_is_valid_code_point_false,
    "System.out.println(Character.isValidCodePoint(0x110000));",
    "false"
);
jt!(
    unicode_to_code_point,
    "System.out.println(Character.toCodePoint((char)0xD83D, (char)0xDE00));",
    "128512"
);
