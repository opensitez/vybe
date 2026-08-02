// vybe-test: kotlin/when_guards/test_when_guarded_math
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun level(v: Int): String = when {
            v * 2 > 10 -> "high"
            v + 1 == 4 -> "four"
            v in 0..2 -> "low"
            else -> "mid"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((level(6)).toString(), "high")
            __check((level(3)).toString(), "four")
            __check((level(1)).toString(), "low")
        }
