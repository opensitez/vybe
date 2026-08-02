// vybe-test: kotlin/spread_arguments/test_spread_with_list_to_array_conversion
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun sum(base: Int, values: IntArray): Int {
            var total = base
            for (v in values) total += v
            return total
        }
        fun main() {
            val items = listOf(1, 2, 3).toIntArray()
            println(sum(1, items))
        }

