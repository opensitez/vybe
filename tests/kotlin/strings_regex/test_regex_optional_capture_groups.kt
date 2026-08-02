// vybe-test: kotlin/strings_regex/test_regex_optional_capture_groups
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\d+)?-(\\w*)")
            val result = pattern.find(" -abc")
            __check((result?.groupValues?.get(1) ?: "missing").toString(), "")
            __check((result?.groupValues?.get(2) ?: "missing").toString(), "abc")
        }
