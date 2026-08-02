// vybe-test: kotlin/scope/test_loop_index_scope_with_outer_name
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            var total = 0
            val label = 5
            for (label in arrayOf(1, 2, 3)) {
                total += label
            }
            println(label)
            println(total)
        }

