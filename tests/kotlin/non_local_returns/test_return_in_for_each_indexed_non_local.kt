// vybe-test: kotlin/non_local_returns/test_return_in_for_each_indexed_non_local
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun findOdd(values: IntArray): Int {
            values.forEachIndexed { index, value ->
                if (index == 2) return value
            }
            return -1
        }

        fun main() {
            println(findOdd(intArrayOf(2, 4, 6, 7, 8)))
        }

