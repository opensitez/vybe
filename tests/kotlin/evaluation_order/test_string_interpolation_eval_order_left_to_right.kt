// vybe-test: kotlin/evaluation_order/test_string_interpolation_eval_order_left_to_right
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun left(): String { order += "L"
return "x" }
            fun right(): String { order += "R"
return "y" }
            val result = "${left()}${right()}"
            __check((result).toString(), "xy")
            __check((order).toString(), "LR")
        }
