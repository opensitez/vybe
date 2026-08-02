// vybe-test: kotlin/functions/test_function_vararg_with_spread_operator
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun sum(label: String, vararg values: Int): String {
            var total = 0
            for (value in values) {
                total += value
            }
            return label + ":" + total
        }

        fun main() {
            val extras = intArrayOf(4, 5)
            println(sum("base", 1, 2, 3, *extras))
            println(sum("empty"))
        }

