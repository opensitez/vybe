// vybe-test: kotlin/strings_regex/test_regex_find_returns_null_when_absent
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("x+")
            val result = pattern.find("abc")
            __check((result == null).toString(), "true")
            val value = pattern.find("abc")?.value ?: "missing"
            __check((value).toString(), "missing")
        }
