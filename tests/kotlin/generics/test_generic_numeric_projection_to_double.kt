// vybe-test: kotlin/generics/test_generic_numeric_projection_to_double
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Number> sumToDouble(values: Array<T>): Double {
            var total = 0.0
            for (value in values) {
                total += value.toDouble()
            }
            return total
        }

        fun main() {
            val ints = arrayOf(1, 2, 3)
            val doubles = arrayOf(1.5, 2.5)
            println(sumToDouble(ints))
            println(sumToDouble(doubles))
        }

