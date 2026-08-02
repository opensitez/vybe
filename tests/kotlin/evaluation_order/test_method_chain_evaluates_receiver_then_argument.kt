// vybe-test: kotlin/evaluation_order/test_method_chain_evaluates_receiver_then_argument
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun arg(v: Int): Int {
                order += "a" + v
                return v
            }
            val out = listOf(1, 2, 3)
                .map { arg(it) }
                .sum()
            __check((out).toString(), "6")
            __check((order).toString(), "a1a2a3")
        }
