// vybe-test: kotlin/non_local_returns/test_non_local_return_with_early_false
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstPositive(values: List<Int>): Int {
            values.forEach {
                if (it < 0) {
                    return 0
                }
            }
            return 1
        }

        fun main() {
            println(firstPositive(listOf(1, 2, -3, 4)))
            println(firstPositive(listOf(1, 2, 3)))
        }

