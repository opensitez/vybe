// vybe-test: kotlin/evaluation_order/test_binary_ops_left_to_right_operands
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun left(): Int { order += "L"
return 1 }
            fun right(): Int { order += "R"
return 2 }
            __check((left() + right()).toString(), "3")
            __check((order).toString(), "LR")
        }
