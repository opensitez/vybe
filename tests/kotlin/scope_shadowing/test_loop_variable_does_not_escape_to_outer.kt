// vybe-test: kotlin/scope_shadowing/test_loop_variable_does_not_escape_to_outer
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun main() {
            var outer = 1
            for (outer in listOf(2, 3, 4)) {
                println(outer)
                break
            }
            println(outer)
        }

