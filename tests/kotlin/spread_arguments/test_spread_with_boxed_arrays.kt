// vybe-test: kotlin/spread_arguments/test_spread_with_boxed_arrays
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun sumAll(values: IntArray): Int {
            var total = 0
            for (v in values) total += v
            return total
        }
        fun main() {
            val a = intArrayOf(1, 2, 3)
            println(sumAll(a))
        }

