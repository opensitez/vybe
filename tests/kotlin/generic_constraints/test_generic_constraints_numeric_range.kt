// vybe-test: kotlin/generic_constraints/test_generic_constraints_numeric_range
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> between(v: T, min: T, max: T): Boolean {
            val n = v.toInt()
            return n in min.toInt()..max.toInt()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((between(3, 1, 5)).toString(), "true")
            __check((between(9, 1, 5)).toString(), "false")
        }
