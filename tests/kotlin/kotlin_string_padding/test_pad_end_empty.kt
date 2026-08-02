// vybe-test: kotlin/kotlin_string_padding/test_pad_end_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_padding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = ""
            __check((s.padStart(2, 'x')).toString(), "xx")
            __check((s.padEnd(2, 'y')).toString(), "yy")
        }
