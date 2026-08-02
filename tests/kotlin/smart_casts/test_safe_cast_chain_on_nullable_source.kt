// vybe-test: kotlin/smart_casts/test_safe_cast_chain_on_nullable_source
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = null
            val text: String? = value as? String
            __check((text ?: "empty").toString(), "empty")
        }
