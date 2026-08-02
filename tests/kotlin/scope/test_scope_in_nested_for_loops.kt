// vybe-test: kotlin/scope/test_scope_in_nested_for_loops
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            var total = 0
            for (row in arrayOf(arrayOf(1, 2), arrayOf(3, 4))) {
                for (col in row) {
                    val value = col * 2
                    total += value
                }
            }
            println(total)
        }

