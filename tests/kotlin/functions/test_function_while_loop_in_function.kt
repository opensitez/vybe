// vybe-test: kotlin/functions/test_function_while_loop_in_function
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun sumUpTo(n: Int): Int {
            var i = 1
            var total = 0
            while (i <= n) {
                total += i
                i += 1
            }
            return total
        }

        fun main() {
            println(sumUpTo(5))
        }

