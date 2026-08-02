// vybe-test: kotlin/nullability/test_elvis_short_circuit_non_null
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name: String? = "Kotlin"
            val display = name ?: "Fallback"
            __check((display).toString(), "Kotlin")
        }
