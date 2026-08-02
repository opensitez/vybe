// vybe-test: kotlin/nullability/test_non_null_variable
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s: String = "Hello"
            __check((s).toString(), "Hello")
        }
