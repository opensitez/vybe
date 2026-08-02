// vybe-test: kotlin/when_guards/test_when_guarded_with_function_call
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun isEven(v: Int): Boolean = v % 2 == 0
        fun label(v: Int): String = when {
            isEven(v) && v > 0 -> "positive-even"
            v > 0 -> "positive"
            else -> "not"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(4)).toString(), "positive-even")
            __check((label(3)).toString(), "positive")
        }
