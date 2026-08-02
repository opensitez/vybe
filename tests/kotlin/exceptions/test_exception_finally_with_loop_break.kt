// vybe-test: kotlin/exceptions/test_exception_finally_with_loop_break
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
    for (value in 1..4) {
        try {
            println(value)
            if (value == 3) {
                break
            }
        } finally {
            println("finally")
        }
    }
    println("done")
}

