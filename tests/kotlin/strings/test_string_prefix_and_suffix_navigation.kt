// vybe-test: kotlin/strings/test_string_prefix_and_suffix_navigation
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "api/v1/resource"
            __check((value.substringBefore("/")).toString(), "api")
            __check((value.substringAfter("/")).toString(), "v1/resource")
            __check((value.substringAfterLast("/")).toString(), "resource")
            __check((value.substringBeforeLast("/")).toString(), "api/v1")
            __check((value.substringBefore("x", "missing")).toString(), "missing")
            __check((value.substringAfter("x", "missing")).toString(), "missing")
        }
