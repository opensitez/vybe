// vybe-test: kotlin/basic/test_nullable_reference_in_basic_flow
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybe: String? = null
            val value: String = maybe ?: "fallback"
            __check((value).toString(), "fallback")
            val explicit: String? = "ok"
            __check((explicit ?: "fallback").toString(), "ok")
        }
