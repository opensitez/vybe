// vybe-test: kotlin/exceptions/test_exception_constructor_failure_is_caught
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

class Exploding {
    init {
        println("init")
        throw Exception("explode")
    }
}

fun main() {
    try {
        val _ = Exploding()
        println("constructed")
    } catch (e: Exception) {
        println("caught")
        println(e.message)
    }
}

