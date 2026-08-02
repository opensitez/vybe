// vybe-test: kotlin/evaluation_order/test_operator_precedence_evaluation_still_left_to_right_for_same_level
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun a(): Int { order += "a"
return 1 }
            fun b(): Int { order += "b"
return 2 }
            fun c(): Int { order += "c"
return 3 }
            val out = a() + b() * c()
            __check((out).toString(), "7")
            __check((order).toString(), "abc")
        }
