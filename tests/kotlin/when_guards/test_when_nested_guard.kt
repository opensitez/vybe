// vybe-test: kotlin/when_guards/test_when_nested_guard
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun check(v: Int): String {
            return when {
                v < 0 -> "neg"
                v in 0..5 -> when (v % 2) {
                    0 -> "small-even"
                    else -> "small-odd"
                }
                else -> "large"
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((check(2)).toString(), "small-even")
            __check((check(3)).toString(), "small-odd")
            __check((check(9)).toString(), "large")
        }
