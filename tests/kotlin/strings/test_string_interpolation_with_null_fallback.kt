// vybe-test: kotlin/strings/test_string_interpolation_with_null_fallback
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nullable: String? = null
            val present: String? = "value"
            __check(("[$nullable]").toString(), "[null]")
            __check(("[${nullable ?: "fallback"}]").toString(), "[fallback]")
            __check(("[${present ?: "fallback"}]").toString(), "[value]")
        }
