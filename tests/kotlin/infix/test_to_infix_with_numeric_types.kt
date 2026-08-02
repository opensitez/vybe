// vybe-test: kotlin/infix/test_to_infix_with_numeric_types
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val combo = 2 to 4.5
            val first = combo.first
            val second = combo.second
            __check((first).toString(), "2")
            __check((second).toString(), "4.5")
        }
