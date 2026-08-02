// vybe-test: kotlin/when_guards/test_when_guarded_char_type
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun toCategory(c: Char): String = when {
            c == 'x' || c == 'y' -> "xy"
            c in 'a'..'f' -> "alpha"
            c.isDigit() -> "digit"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((toCategory('x')).toString(), "xy")
            __check((toCategory('b')).toString(), "alpha")
            __check((toCategory('7')).toString(), "digit")
        }
