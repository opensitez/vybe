// vybe-test: kotlin/scope/test_scope_shadowed_loop_variable_stays_local_to_iteration_body
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            val values = arrayOf(1, 2, 3)
            var sum = 0

            for (value in values) {
                run {
                    val value = value * 10
                    sum += value
                }
            }

            println(sum)
        }

