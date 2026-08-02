// vybe-test: kotlin/try_catch_flow/test_try_no_exception
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                println("ok")
            } catch (e: Exception) {
                println("fail")
            }
        }

