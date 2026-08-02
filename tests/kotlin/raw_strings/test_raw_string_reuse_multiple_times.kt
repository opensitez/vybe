// vybe-test: kotlin/raw_strings/test_raw_string_reuse_multiple_times
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun make(base: String): String = """[$base]"""
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = make("x") + make("y")
            __check((x).toString(), "[x][y]")
        }
