// vybe-test: kotlin/try_finally/test_nested_try_finally_with_catch
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        try {
            try {
                throw RuntimeException("x")
            } catch (e: RuntimeException) {
                println("inner")
            } finally {
                println("inner-finally")
            }
        } catch (e: Exception) {
            println("outer")
        } finally {
            println("outer-finally")
        }
    }

