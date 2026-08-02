// vybe-test: kotlin/non_local_returns/test_anonymous_function_return_is_local
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun total(values: List<Int>): Int {
            var total = 0
            values.forEach(fun(value: Int) {
                if (value < 0) {
                    return
                }
                total += value
            })
            return total
        }

        fun main() {
            println(total(listOf(-1, 2, -3, 4)))
        }

