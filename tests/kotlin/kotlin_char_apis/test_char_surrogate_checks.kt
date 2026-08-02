// vybe-test: kotlin/kotlin_char_apis/test_char_surrogate_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = '\uD83D'
            __check((c.isSurrogate()).toString(), "true")
            __check((c.isHighSurrogate()).toString(), "true")
            val d = '\uDE00'
            __check((d.isLowSurrogate()).toString(), "true")
            __check((d.isSurrogate()).toString(), "true")
        }
