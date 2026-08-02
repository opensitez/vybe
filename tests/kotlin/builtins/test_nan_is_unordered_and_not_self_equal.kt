// vybe-test: kotlin/builtins/test_nan_is_unordered_and_not_self_equal
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0.0 / 0.0
            __check((value.isNaN()).toString(), "true")
            __check((value == value).toString(), "false")
        }
