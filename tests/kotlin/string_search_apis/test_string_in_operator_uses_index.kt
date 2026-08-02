// vybe-test: kotlin/string_search_apis/test_string_in_operator_uses_index
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin"
            __check(('t' in text).toString(), "true")
            __check(('x' in text).toString(), "false")
            __check((3 in 1..5).toString(), "true")
        }
