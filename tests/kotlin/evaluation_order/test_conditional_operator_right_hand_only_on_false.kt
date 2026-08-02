// vybe-test: kotlin/evaluation_order/test_conditional_operator_right_hand_only_on_false
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            val value = if (true) { log += "t"
1 } else { log += "f"
2 }
            __check((value).toString(), "1")
            __check((log).toString(), "t")
        }
