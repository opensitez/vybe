// vybe-test: kotlin/basic/test_boolean_operator_short_circuit_true_or
// origin: languages/kotlin/tests/kotlin/test_basic.rs

var called = 0

        fun sideEffect(): Boolean {
            called += 1
            return false
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || sideEffect()).toString(), "true")
            __check((called).toString(), "0")
            __check((false || sideEffect()).toString(), "false")
            __check((called).toString(), "1")
        }
