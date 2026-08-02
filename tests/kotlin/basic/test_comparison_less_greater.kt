// vybe-test: kotlin/basic/test_comparison_less_greater
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((5 < 10).toString(), "true")
            __check((15 > 10).toString(), "true")
        }
