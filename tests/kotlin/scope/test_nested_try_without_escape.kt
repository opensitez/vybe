// vybe-test: kotlin/scope/test_nested_try_without_escape
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun main() {
            val value = 1
            try {
                val value = "inner"
                println(value)
            } catch (e: Exception) {
                println("err")
            }
            println(value)
        }

