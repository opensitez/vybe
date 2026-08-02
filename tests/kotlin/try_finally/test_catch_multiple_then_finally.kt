// vybe-test: kotlin/try_finally/test_catch_multiple_then_finally
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        try {
            throw IllegalArgumentException()
        } catch (e: IllegalStateException) {
            println("state")
        } catch (e: IllegalArgumentException) {
            println("arg")
        } finally {
            println("done")
        }
    }

