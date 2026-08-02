// vybe-test: kotlin/evaluation_order/test_nested_calls_evaluate_outer_before_inner_return
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun combine(x: Int, y: Int): Int = x * 10 + y
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun f(): Int { order += "f"
return 1 }
            fun g(x: Int): Int { order += "g"
return x + 1 }
            fun h(x: Int): Int { order += "h"
return x + 2 }
            val out = combine(g(f()), h(g(3)))
            __check((out).toString(), "26")
            __check((order).toString(), "fghg")
        }
