// vybe-test: kotlin/local_returns/test_local_return_nonlocal_for_each
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun search(values: List<Int>): String {
            for (v in values) {
                if (v == 3) return "found"
            }
            return "none"
        }
        fun main() {
            println(search(listOf(1, 2, 3)))
            println(search(listOf(1, 2)))
        }

