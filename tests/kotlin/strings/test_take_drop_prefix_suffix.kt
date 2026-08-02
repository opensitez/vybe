// vybe-test: kotlin/strings/test_take_drop_prefix_suffix
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "abcdef"
            __check((word.take(3)).toString(), "abc")
            __check((word.drop(3)).toString(), "def")
            __check((word.takeLast(2)).toString(), "ef")
            __check((word.dropLast(4)).toString(), "ab")
        }
