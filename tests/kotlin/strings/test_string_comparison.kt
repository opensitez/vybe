// vybe-test: kotlin/strings/test_string_comparison
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ab" < "ac").toString(), "true")
            __check(("ab" == "ab").toString(), "true")
            __check(("ab" != "BA").toString(), "true")
        }
