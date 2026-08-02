// vybe-test: kotlin/infix/test_custom_contains_operator_infix_style
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Window(val min: Int, val max: Int) {
            operator fun contains(value: Int): Boolean = value in min..max
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val active = Window(1, 5)
            __check((3 in active).toString(), "true")
            __check((7 in active).toString(), "false")
            __check((8 !in active).toString(), "true")
        }
