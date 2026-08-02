// vybe-test: kotlin/exceptions/test_throw_and_catch_specific
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
            try {
                throw Exception("boom")
            } catch (e: IllegalArgumentException) {
                println("arg")
            } catch (e: Exception) {
                println("general")
            } finally {
                println("complete")
            }
        }

