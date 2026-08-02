// vybe-test: kotlin/infix/test_infix_contains_from_extension
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Window(val values: Set<Int>) {
            operator fun contains(value: Int): Boolean = values.contains(value)
        }

        infix fun Window.has(value: Int): Boolean = value in this

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val setWindow = Window(setOf(2, 4, 6))
            __check((setWindow has 4).toString(), "true")
            __check((setWindow has 3).toString(), "false")
        }
