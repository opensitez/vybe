// vybe-test: kotlin/scope/test_scope_with_shadowed_loop_variables
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            val i = 100
            var output = 0
            for (i in arrayOf(1, 2, 3)) {
                output += i
            }
            println(i)
            println(output)
        }

