// vybe-test: kotlin/scope/test_scope_in_for_each_lambda_and_outer_capture
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            var total = 0
            val source = listOf(1, 2, 3)
            source.forEach {
                val transformed = it * 2
                total += transformed
            }
            println(total)
            println(source.size)
        }

