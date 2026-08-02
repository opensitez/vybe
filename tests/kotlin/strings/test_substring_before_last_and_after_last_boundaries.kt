// vybe-test: kotlin/strings/test_substring_before_last_and_after_last_boundaries
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a/b/c"
            __check((value.substringBeforeLast("/")).toString(), "a/b")
            __check((value.substringAfterLast("/")).toString(), "c")
            __check(("nodelim".substringAfterLast("/", "none")).toString(), "none")
            __check(("x/y/".substringAfterLast("/")).toString(), "")
        }
