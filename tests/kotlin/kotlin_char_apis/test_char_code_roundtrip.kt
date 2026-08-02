// vybe-test: kotlin/kotlin_char_apis/test_char_code_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'Z'
            val code = c.code
            __check((code).toString(), "90")
            __check((code.toChar()).toString(), "Z")
            __check(('Ω'.code).toString(), "937")
            __check(('Ω'.code.toChar()).toString(), "Ω")
        }
