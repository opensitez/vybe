// vybe-test: kotlin/exceptions/test_exception_nested_function_throwing
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun inner(): Int {
    throw Exception("inner")
}

fun outer() {
    try {
        inner()
    } catch (e: Exception) {
        __check(("outer").toString(), "outer")
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    outer()
}
