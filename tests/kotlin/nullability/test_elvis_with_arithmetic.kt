// vybe-test: kotlin/nullability/test_elvis_with_arithmetic
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count: Int? = null
            val total = (count ?: 0) + 10
            __check((total).toString(), "10")
        }
