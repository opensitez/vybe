// vybe-test: kotlin/exceptions/test_exception_try_with_continue_in_catch_then_finally
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
    for (value in 1..4) {
        try {
            if (value == 2) {
                throw Exception("bad")
            }
            println(value)
            continue
        } catch (e: Exception) {
            println("caught")
            continue
        } finally {
            println("finally")
        }
    }
    println("done")
}

