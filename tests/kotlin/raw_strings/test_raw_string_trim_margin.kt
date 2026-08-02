// vybe-test: kotlin/raw_strings/test_raw_string_trim_margin
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                |one
                |two
                |three
            """.trimMargin()
            __check((text).toString(), "one\ntwo\nthree")
        }
