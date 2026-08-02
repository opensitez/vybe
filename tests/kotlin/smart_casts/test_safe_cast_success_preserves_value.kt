// vybe-test: kotlin/smart_casts/test_safe_cast_success_preserves_value
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "abc"
            val text: String? = value as? String
            __check((text ?: "missing").toString(), "abc")
        }
