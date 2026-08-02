// vybe-test: kotlin/exceptions/test_try_catch_multiple_handlers
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: IllegalArgumentException) {
                println("arg")
            } catch (e: Exception) {
                println("general")
            }
        }

