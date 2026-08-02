// vybe-test: kotlin/strings_regex/test_regex_multiple_captures_mapping
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\d)(\\w)(\\w)")
            val result = pattern.find("9ab")
            val mapped = result?.groups?.let {
                "${it[1]?.value}-${it[2]?.value}-${it[3]?.value}"
            } ?: "none"
            __check((mapped).toString(), "9-a-b")
            __check((pattern.find("abc") == null).toString(), "true")
        }
