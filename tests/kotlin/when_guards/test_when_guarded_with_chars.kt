// vybe-test: kotlin/when_guards/test_when_guarded_with_chars
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(c: Char): String = when {
            c in 'a'..'f' -> "alpha"
            c in 'g'..'m' -> "middle"
            c in 'n'..'z' -> "late"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label('c')).toString(), "alpha")
            __check((label('k')).toString(), "middle")
            __check((label('z')).toString(), "late")
        }
