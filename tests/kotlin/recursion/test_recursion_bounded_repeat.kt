// vybe-test: kotlin/recursion/test_recursion_bounded_repeat
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun repeatText(value: String, count: Int): String = if (count <= 0) "" else value + " " + repeatText(value, count - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((repeatText("x", 2)).toString(), "x x ")
        }
