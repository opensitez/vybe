// vybe-test: kotlin/evaluation_order/test_constructor_arg_order
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

class Tracker {
            constructor(a: Int, b: Int) {
                __check((a).toString(), "1")
                __check((b).toString(), "2")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun first(): Int { log += "1"
return 1 }
            fun second(): Int { log += "2"
return 2 }
            Tracker(first(), second())
            __check((log).toString(), "12")
        }
