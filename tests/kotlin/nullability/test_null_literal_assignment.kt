// vybe-test: kotlin/nullability/test_null_literal_assignment
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s: String? = null
            __check((s).toString(), "null")
        }
