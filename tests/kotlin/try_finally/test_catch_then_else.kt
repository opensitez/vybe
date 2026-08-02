// vybe-test: kotlin/try_finally/test_catch_then_else
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        try {
            println("ok")
        } catch (e: Exception) {
            println("bad")
        } finally {
            println("fin")
        }
    }

