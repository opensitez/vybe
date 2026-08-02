// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_constraint_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

class NumericBox<T>(val value: T) where T : Number, T : Comparable<T>

        fun <T> maxBox(a: NumericBox<T>, b: NumericBox<T>): T where T : Number, T : Comparable<T> {
            return if (a.value >= b.value) a.value else b.value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = NumericBox(3)
            val b = NumericBox(7)
            __check((maxBox(a, b)).toString(), "7")
        }
