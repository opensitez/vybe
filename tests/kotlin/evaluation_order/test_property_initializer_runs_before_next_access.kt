// vybe-test: kotlin/evaluation_order/test_property_initializer_runs_before_next_access
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

var order = ""
        val one = run {
            order += "1"
            1
        }
        val two = run {
            order += "2"
            2
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((one + two).toString(), "3")
            __check((order).toString(), "12")
        }
