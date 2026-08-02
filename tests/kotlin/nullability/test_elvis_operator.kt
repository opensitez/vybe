// vybe-test: kotlin/nullability/test_elvis_operator
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name: String? = null
            val display = name ?: "default"
            __check((display).toString(), "default")
        }
