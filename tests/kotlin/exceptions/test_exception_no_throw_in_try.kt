// vybe-test: kotlin/exceptions/test_exception_no_throw_in_try
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
    try {
        println("safe")
    } catch (e: Exception) {
        println("caught")
    } finally {
        println("done")
    }
}

