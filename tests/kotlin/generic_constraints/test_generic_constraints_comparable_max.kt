// vybe-test: kotlin/generic_constraints/test_generic_constraints_comparable_max
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> maxValue(a: T, b: T): T where T : Comparable<T> {
            return if (a >= b) a else b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxValue(9, 10)).toString(), "10")
            __check((maxValue("m", "n")).toString(), "n")
        }
