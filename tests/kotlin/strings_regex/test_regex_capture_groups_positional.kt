// vybe-test: kotlin/strings_regex/test_regex_capture_groups_positional
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\d{2})-(\\w+)")
            val result = pattern.find("42-abc")
            __check((result?.destructured?.component1() ?: "none").toString(), "42")
            __check((result?.destructured?.component2() ?: "none").toString(), "abc")
            __check((result?.groupValues?.size ?: 0).toString(), "3")
        }
