// vybe-test: kotlin/scope_shadowing/test_for_each_lambda_parameter_shadowing_outer
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun main() {
            val value = "outer"
            val values = listOf("a", "b")
            values.forEach { value ->
                println(value)
            }
            println(value)
        }

