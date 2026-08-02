// vybe-test: kotlin/evaluation_order/test_when_subject_and_guard_order
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            val v = run {
                order += "s"
                7
            }
            val out = when {
                (order += "g").isNotEmpty() && v > 0 -> {
                    order += "t"
                    "yes"
                }
                else -> {
                    order += "f"
                    "no"
                }
            }
            __check((out).toString(), "yes")
            __check((order).toString(), "sgt")
        }
