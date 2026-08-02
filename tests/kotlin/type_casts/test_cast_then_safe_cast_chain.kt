// vybe-test: kotlin/type_casts/test_cast_then_safe_cast_chain
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val source: Any? = "x"
            val direct = source as String
            val safe = (source as? String) ?: "fallback"
            val failed = source as? Int
            __check((direct + ":" + safe).toString(), "x:x")
    __check((failed == null).toString(), "true")
}
