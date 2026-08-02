// vybe-test: kotlin/kotlin_infix_keywords/test_infix_is_operator
// origin: languages/kotlin/tests/kotlin/test_kotlin_infix_keywords.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x: Any = "x"
            val y = "x"
            __check(((x is String).toString()).toString(), "true")
            __check(((x !is Int).toString()).toString(), "true")
            __check(((y is String).toString()).toString(), "true")
        }
