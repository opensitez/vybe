// vybe-test: kotlin/scope/test_while_scope_and_mutation
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            var total = 0
            var index = 0
            while (index < 3) {
                val step = index + 1
                total += step
                index += 1
            }
            println(total)
        }

