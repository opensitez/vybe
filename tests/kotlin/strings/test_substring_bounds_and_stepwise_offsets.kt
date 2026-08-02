// vybe-test: kotlin/strings/test_substring_bounds_and_stepwise_offsets
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "compiler"
            __check((word.substring(0, 3)).toString(), "com")
            __check((word.substring(3)).toString(), "iler")
            __check((word.substring(word.length - 2)).toString(), "er")
        }
