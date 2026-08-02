// vybe-test: kotlin/smart_casts/test_safe_cast_returns_null_when_type_mismatch
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 7
            val text: String? = value as? String
            __check((text == null).toString(), "true")
        }
