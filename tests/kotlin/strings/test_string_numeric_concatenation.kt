// vybe-test: kotlin/strings/test_string_numeric_concatenation
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count = 2
            __check(("count=" + count).toString(), "count=2")
            __check(("next=${count + 1}").toString(), "next=3")
        }
