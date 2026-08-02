// vybe-test: kotlin/strings/test_nullable_string_is_null_or_empty
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val missing: String? = null
            val empty: String? = ""
            __check((missing.isNullOrEmpty()).toString(), "true")
            __check((empty.isNullOrEmpty()).toString(), "true")
            __check((("abc").isNullOrEmpty()).toString(), "false")
        }
