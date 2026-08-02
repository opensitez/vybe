// vybe-test: kotlin/exceptions/test_try_without_exception
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
            try {
                println("normal flow")
            } catch (e: Exception) {
                println("error")
            }
        }

