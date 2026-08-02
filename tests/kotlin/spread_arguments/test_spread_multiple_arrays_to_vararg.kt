// vybe-test: kotlin/spread_arguments/test_spread_multiple_arrays_to_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(values: IntArray): Int {
            var total = 0
            for (v in values) total += v
            return total
        }
        fun sum(prefix: String, vararg values: Int): Int {
            return prefix.length + values.sum()
        }
        fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(3, 4)
            println(sum("x", *a, *b))
        }

