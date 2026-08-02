// vybe-test: kotlin/non_local_returns/test_non_local_return_from_take_if
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstMatching(values: List<Int>): Int {
            values.takeIf { it.isNotEmpty() }?.forEach {
                if (it == 2) return it
            }
            return -1
        }

        fun main() {
            println(firstMatching(listOf(1, 2, 3)))
        }

