// vybe-test: kotlin/kotlin_range_edge_conditions/test_char_range_iteration
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = ('a'..'d').joinToString("")
            __check((text).toString(), "abcd")
        }
