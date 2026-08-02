// vybe-test: kotlin/infix/test_to_infix_nested_destructure
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (left, right) = "left" to "right"
            __check((left + "," + right).toString(), "left,right")
        }
