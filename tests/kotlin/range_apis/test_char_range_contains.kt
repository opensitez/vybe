// vybe-test: kotlin/range_apis/test_char_range_contains
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 'a'..'d'
            __check((r.first).toString(), "a")
            __check((r.last).toString(), "d")
            __check(('c' in r).toString(), "true")
            __check(('x' in r).toString(), "false")
        }
