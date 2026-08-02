// vybe-test: kotlin/local_functions/test_local_function_in_while_like_rewrite
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun main() {
            fun next(base: Int): Int = base + 1
            var i = 0
            var sum = 0
            while (i < 4) {
                sum += next(i)
                i++
            }
            println(sum)
        }

