// vybe-test: kotlin/try_finally/test_finally_and_exception_type
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        try {
            val x = 1 / 0
            println(x)
        } catch (e: ArithmeticException) {
            println("arith")
        } finally {
            println("done")
        }
    }

