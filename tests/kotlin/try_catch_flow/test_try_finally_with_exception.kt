// vybe-test: kotlin/try_catch_flow/test_try_finally_with_exception
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                val x = 1 / 0
                println(x)
            } finally {
                println("finally")
            }
        }

