// vybe-test: kotlin/kotlin_string_padding/test_pad_start_already_wide
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_padding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abc"
            __check((s.padStart(2)).toString(), "abc")
        }
