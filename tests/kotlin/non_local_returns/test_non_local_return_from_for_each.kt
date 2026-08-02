// vybe-test: kotlin/non_local_returns/test_non_local_return_from_for_each
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstPositive(values: List<Int>): Int {
            values.forEach {
                if (it > 0) return it
            }
            return -1
        }

        fun main() {
            println(firstPositive(listOf(-2, 0, 3, 4)) )
        }

