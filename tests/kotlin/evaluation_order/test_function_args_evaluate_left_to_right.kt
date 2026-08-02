// vybe-test: kotlin/evaluation_order/test_function_args_evaluate_left_to_right
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun lhs(): Int { order += "L"
return 1 }
            fun mid(v: Int): Int { order += "M"
return v + 1 }
            fun rhs(v: Int): Int { order += "R"
return v + 2 }
            fun f(a: Int, b: Int, c: Int) {
                __check((a).toString(), "1")
                __check((b).toString(), "1")
                __check((c).toString(), "2")
            }
            f(lhs(), mid(0), rhs(0))
            __check((order).toString(), "LMR")
        }
