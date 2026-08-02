// vybe-test: kotlin/local_functions/test_local_function_in_for_loop_body
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun main() {
            val values = listOf(1, 2, 3)
            var total = 0
            fun add(v: Int) {
                total += v
            }
            for (value in values) {
                add(value)
            }
            println(total)
        }

