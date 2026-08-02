// vybe-test: kotlin/evaluation_order/test_multiple_expressions_in_block_preserve_order
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val order = run {
                val out = StringBuilder()
                out.append("a")
                out.append("b")
                out.append("c")
                out.toString()
            }
            __check((order).toString(), "abc")
        }
