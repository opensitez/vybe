// vybe-test: kotlin/evaluation_order/test_not_operator_takes_single_evaluation
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun flag(): Boolean {
                order += "f"
                return false
            }
            __check((!flag()).toString(), "true")
            __check((order).toString(), "f")
        }
