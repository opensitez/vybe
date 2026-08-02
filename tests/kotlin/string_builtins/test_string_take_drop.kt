// vybe-test: kotlin/string_builtins/test_string_take_drop
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "abcdef"
            __check((text.take(2)).toString(), "ab")
            __check((text.drop(3)).toString(), "def")
            __check((text.dropLast(2)).toString(), "abcd")
        }
