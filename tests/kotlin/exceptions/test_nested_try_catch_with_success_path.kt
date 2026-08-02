// vybe-test: kotlin/exceptions/test_nested_try_catch_with_success_path
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
            try {
                try {
                    println("inner")
                } catch (e: Exception) {
                    println("should not")
                } finally {
                    println("inner done")
                }
            } catch (e: Exception) {
                println("outer catch")
            } finally {
                println("outer done")
            }
        }

