// vybe-test: kotlin/exceptions/test_exception_with_else_path
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun main() {
    try {
        throw Exception()
    } catch (e: IllegalArgumentException) {
        println("arg")
    } catch (e: Exception) {
        println("general")
    }
}

