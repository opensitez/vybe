// vybe-test: kotlin/strings/test_string_escape_sequence
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val quoted = "He said \"Kotlin\""
            val path = "C:\\temp\\out"
            __check((quoted).toString(), "He said \"Kotlin\"")
            __check((path).toString(), "C:\\temp\\out")
        }
