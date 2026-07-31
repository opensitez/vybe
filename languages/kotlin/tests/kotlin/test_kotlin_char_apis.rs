kotlin_run_cases! {
    test_char_classification_letters_digits_whitespace => (r##"
        fun main() {
            println('A'.isLetter())
            println('3'.isDigit())
            println(' '.isWhitespace())
            println('a'.isLetterOrDigit())
            println('?'.isLetterOrDigit())
        }
    "##, &[
        "true",
        "true",
        "true",
        "true",
        "false",
    ]),
    test_char_case_predicates => (r##"
        fun main() {
            println('a'.isUpperCase())
            println('A'.isUpperCase())
            println('a'.isLowerCase())
            println('A'.isLowerCase())
            println('9'.isUpperCase())
        }
    "##, &[
        "false",
        "true",
        "true",
        "false",
        "false",
    ]),
    test_char_surrogate_checks => (r##"
        fun main() {
            val c = '\uD83D'
            println(c.isSurrogate())
            println(c.isHighSurrogate())
            val d = '\uDE00'
            println(d.isLowSurrogate())
            println(d.isSurrogate())
        }
    "##, &[
        "true",
        "true",
        "true",
        "true",
    ]),
    test_char_control_and_defined => (r##"
        fun main() {
            println('\u0000'.isISOControl())
            println('A'.isISOControl())
            println('A'.isDefined())
            println('\uFFFF'.isDefined())
        }
    "##, &[
        "true",
        "false",
        "true",
        "false",
    ]),
    test_char_identifier_checks => (r##"
        fun main() {
            println('a'.isLetter())
            println(''.isLetter())
            println('7'.isDigit())
            println('_'.isIdentifierStart())
            println('x'.isIdentifierPart())
            println(' '.isIdentifierPart())
        }
    "##, &[
        "true",
        "false",
        "true",
        "true",
        "true",
        "false",
    ]),
    test_char_case_transform => (r##"
        fun main() {
            val c = 'a'
            println(c.uppercaseChar())
            println(c.lowercaseChar())
            println('ß'.uppercase())
            println('A'.uppercase().toString())
        }
    "##, &[
        "A",
        "a",
        "SS",
        "A",
    ]),
    test_char_titlecase => (r##"
        fun main() {
            val c = 'a'
            println(c.titlecase())
            val d = 'ǈ'
            println(d.isTitleCase())
            println('A'.isTitleCase())
        }
    "##, &[
        "A",
        "false",
        "false",
    ]),
    test_char_code_roundtrip => (r##"
        fun main() {
            val c = 'Z'
            val code = c.code
            println(code)
            println(code.toChar())
            println('Ω'.code)
            println('Ω'.code.toChar())
        }
    "##, &[
        "90",
        "Z",
        "937",
        "Ω",
    ]),
    test_char_sequence_navigation => (r##"
        fun main() {
            val s = "Abc"
            println(s.first())
            println(s.last())
            println(s.elementAt(1))
            println(s[2])
        }
    "##, &[
        "A",
        "c",
        "b",
        "c",
    ]),
    test_char_to_digit_and_back => (r##"
        fun main() {
            println('9'.digitToInt())
            println('a'.digitToIntOrNull())
            println('a'.digitToInt(16))
            println('f'.digitToIntOrNull(16))
        }
    "##, &[
        "9",
        "null",
        "10",
        "15",
    ]),
}
