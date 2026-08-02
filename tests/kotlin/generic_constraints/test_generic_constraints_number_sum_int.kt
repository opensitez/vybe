// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_sum_int
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> add(a: T, b: T): Int = a.toInt() + b.toInt()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((add(1, 2)).toString(), "3")
        }
