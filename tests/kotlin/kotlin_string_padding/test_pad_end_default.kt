// vybe-test: kotlin/kotlin_string_padding/test_pad_end_default
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_padding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "5"
            __check((s.padEnd(3)).toString(), "5  ")
        }
