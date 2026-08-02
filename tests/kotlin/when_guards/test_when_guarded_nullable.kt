// vybe-test: kotlin/when_guards/test_when_guarded_nullable
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(v: Int?): String = when {
            v == null -> "null"
            v > 3 -> "big"
            else -> "small"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(null)).toString(), "null")
            __check((label(2)).toString(), "small")
        }
