// vybe-test: kotlin/kotlin_infix_keywords/test_infix_to_pair
// origin: languages/kotlin/tests/kotlin/test_kotlin_infix_keywords.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = 1 to 2
            __check((pair.first).toString(), "1")
            __check((pair.second).toString(), "2")
        }
