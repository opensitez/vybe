// vybe-test: kotlin/string_builtins/test_string_repeat_and_padding
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "x"
            __check((text.repeat(3)).toString(), "xxx")
            __check(("1".padStart(3, '0')).toString(), "001")
            __check(("7".padEnd(3, '0')).toString(), "700")
        }
