// vybe-test: kotlin/strings/test_string_chunked_and_windowed_views
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "abcdef"
            __check((value.chunked(2).joinToString("|")).toString(), "ab|cd|ef")
            val windows = value.windowed(3, 2)
            __check((windows.joinToString("|")).toString(), "abc|cde")
        }
