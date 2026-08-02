// vybe-test: kotlin/boolean_logic/test_boolean_arithmetic_operators_and_assignment_mix
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var a = true
            var b = false
            a = a && !b
            b = a || !b
            __check((a).toString(), "true")
            __check((b).toString(), "true")
            a = a.xor(b)
            __check((a).toString(), "false")
        }
