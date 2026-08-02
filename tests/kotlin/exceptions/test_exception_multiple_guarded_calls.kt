// vybe-test: kotlin/exceptions/test_exception_multiple_guarded_calls
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun fail(v: Int) {
    if (v == 1) {
        throw Exception("x")
    }
}

fun main() {
    for (v in arrayOf(0, 1, 2)) {
        try {
            fail(v)
            println(v)
        } catch (e: Exception) {
            println("err")
        }
    }
}

