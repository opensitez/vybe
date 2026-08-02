// vybe-test: kotlin/functions/test_function_local_var_mutation
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun accum(n: Int): Int {
            var sum = 0
            var i = 1
            while (i <= n) {
                sum += i
                i += 1
            }
            return sum
        }

        fun main() {
            println(accum(4))
        }

