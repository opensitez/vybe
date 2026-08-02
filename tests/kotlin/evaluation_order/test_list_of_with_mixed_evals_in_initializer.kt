// vybe-test: kotlin/evaluation_order/test_list_of_with_mixed_evals_in_initializer
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun a(): Int { order += "A"
return 1 }
            fun b(): Int { order += "B"
return 2 }
            fun c(): Int { order += "C"
return 3 }
            val values = listOf(a(), b(), c())
            __check((values.joinToString(",")).toString(), "1,2,3")
            __check((order).toString(), "ABC")
        }
