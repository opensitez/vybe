// vybe-test: kotlin/basic/test_comparison_chain
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = 82
            __check((score > 80 && score < 90).toString(), "true")
            __check((score < 50 || score == 82).toString(), "true")
        }
