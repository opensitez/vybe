// vybe-test: kotlin/evaluation_order/test_ternary_like_when_not_available
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = if (true) "yes" else "no"
            __check((value).toString(), "yes")
        }
