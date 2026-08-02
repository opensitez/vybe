// vybe-test: kotlin/exceptions/test_exception_finally_with_continue_path
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
    for (v in 1..3) {
        try {
            println(v)
        } finally {
            println("tick")
        }
    }
}

