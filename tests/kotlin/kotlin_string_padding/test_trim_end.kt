// vybe-test: kotlin/kotlin_string_padding/test_trim_end
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_padding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "  abc  "
            __check((s.trimEnd()).toString(), "  abc")
        }
