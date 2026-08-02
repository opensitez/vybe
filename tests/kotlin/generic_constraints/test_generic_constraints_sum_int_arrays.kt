// vybe-test: kotlin/generic_constraints/test_generic_constraints_sum_int_arrays
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> sumArray(v: Array<T>): Double {
            var total = 0.0
            for (item in v) total += item.toDouble()
            return total
        }
        fun main() {
            println(sumArray(arrayOf(1, 2, 3)))
        }

