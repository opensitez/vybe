// vybe-test: kotlin/functions/test_function_varargs
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun sumAll(vararg values: Int): Int {
            var total = 0
            var i = 0
            while (i < 4) {
                total += values[i]
                i += 1
            }
            return total
        }

        fun main() {
            println(sumAll(1, 2, 3, 4))
        }

