// vybe-test: kotlin/basic/test_comparison_less_equal_greater_equal
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10 <= 10).toString(), "true")
            __check((10 >= 10).toString(), "true")
        }
