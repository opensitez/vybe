// vybe-test: kotlin/evaluation_order/test_default_arguments_evaluate_when_omitted
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var order = ""
            fun value(prefix: String = "d"): String {
                order += "v"
                return prefix
            }
            fun report(a: String = value("a"), b: String = value("b")) {
                __check((a).toString(), "x")
                __check((b).toString(), "b")
            }
            report("x")
            __check((order).toString(), "v")
        }
