// vybe-test: kotlin/kotlin_string_find_contains/test_find_prefix
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcabc"
            __check((s.indexOf('a', startIndex = 1).toString()).toString(), "3")
            __check((s.substringAfter("ab").toString()).toString(), "abc")
        }
