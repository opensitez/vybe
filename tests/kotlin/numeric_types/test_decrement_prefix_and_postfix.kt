// vybe-test: kotlin/numeric_types/test_decrement_prefix_and_postfix
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 10
            __check((--value).toString(), "9")
            __check((value).toString(), "9")
            __check((value--).toString(), "9")
            __check((value).toString(), "8")
        }
