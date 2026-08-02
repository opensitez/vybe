// vybe-test: kotlin/when_guards/test_when_with_guard_first
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(x: Int): String = when {
            x < 0 -> "neg"
            x == 0 -> "zero"
            x % 2 == 0 -> "even"
            else -> "odd"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(-1)).toString(), "neg")
            __check((label(4)).toString(), "even")
            __check((label(3)).toString(), "odd")
        }
