// vybe-test: kotlin/try_catch_flow/test_try_shadowed_variable
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            val x = 1
            try {
                val x = 5
                println(x)
            } catch (e: Exception) {
                println(0)
            }
            println(x)
        }

