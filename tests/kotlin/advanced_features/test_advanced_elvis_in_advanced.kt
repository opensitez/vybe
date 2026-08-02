// vybe-test: kotlin/advanced_features/test_advanced_elvis_in_advanced
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val text: String? = null
__check((text ?: "none").toString(), "none") }
