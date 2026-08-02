// vybe-test: kotlin/when_guards/test_when_guarded_char
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'z'
            val out = when {
                c in 'a'..'m' -> "first"
                c in 'n'..'z' -> "last"
                else -> "other"
            }
            __check((out).toString(), "last")
        }
