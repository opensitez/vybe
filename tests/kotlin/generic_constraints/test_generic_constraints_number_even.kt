// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_even
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> isEven(v: T): Boolean = v.toInt() % 2 == 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isEven(4)).toString(), "true")
            __check((isEven(5)).toString(), "false")
        }
