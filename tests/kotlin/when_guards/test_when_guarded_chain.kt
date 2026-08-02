// vybe-test: kotlin/when_guards/test_when_guarded_chain
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(a: Int, b: Int): String = when {
            a == 0 || b == 0 -> "zero"
            a == b -> "equal"
            a + b == 10 -> "ten"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(0, 5)).toString(), "zero")
            __check((label(5, 5)).toString(), "equal")
            __check((label(3, 7)).toString(), "ten")
        }
