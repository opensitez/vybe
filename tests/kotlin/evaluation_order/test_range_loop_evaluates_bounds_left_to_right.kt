// vybe-test: kotlin/evaluation_order/test_range_loop_evaluates_bounds_left_to_right
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            val start = run { order += "s"
1 }
            val end = run { order += "e"
3 }
            val values = (start..end).toList().joinToString(",")
            __check((values).toString(), "1,2,3")
            __check((order).toString(), "se")
        }
