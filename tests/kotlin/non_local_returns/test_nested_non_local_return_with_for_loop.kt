// vybe-test: kotlin/non_local_returns/test_nested_non_local_return_with_for_loop
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstOdd(values: List<Int>): Int {
            values.forEach {
                run {
                    if (it % 2 == 1) {
                        return it
                    }
                }
            }
            return -1
        }

        fun main() {
            println(firstOdd(listOf(2, 4, 6, 9, 10)))
        }

