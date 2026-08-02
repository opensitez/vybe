// vybe-test: kotlin/kotlin_char_predicates/test_char_case_conversion
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_predicates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'b'
            __check((c.toUpperCase()).toString(), "B")
            __check((c.toLowerCase()).toString(), "b")
        }
