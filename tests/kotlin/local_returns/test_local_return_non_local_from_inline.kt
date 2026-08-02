// vybe-test: kotlin/local_returns/test_local_return_non_local_from_inline
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun sum(values: List<Int>): Int {
            var s = 0
            values.forEach {
                if (it < 0) return 0
                s += it
            }
            return s
        }
        fun main() {
            println(sum(listOf(1, 2, 3)))
            println(sum(listOf(-1, 2)))
        }

