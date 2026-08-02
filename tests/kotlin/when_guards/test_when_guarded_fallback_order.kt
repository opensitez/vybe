// vybe-test: kotlin/when_guards/test_when_guarded_fallback_order
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(v: Int): String = when {
            v < 0 -> "neg"
            v % 2 == 0 -> "even"
            else -> "odd"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(-2)).toString(), "neg")
            __check((label(3)).toString(), "odd")
        }
