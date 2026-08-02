// vybe-test: kotlin/operators/test_logical_operators_preserve_left_to_right_evaluation_for_side_effects
// origin: languages/kotlin/tests/kotlin/test_operators.rs

var trace = ""
        fun hit(label: String, value: Boolean): Boolean {
            trace += label
            return value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((hit("a", false) && hit("b", true)).toString(), "false")
            __check((trace).toString(), "a")
            trace = ""
            __check((hit("c", true) || hit("d", false)).toString(), "true")
            __check((trace).toString(), "c")
        }
