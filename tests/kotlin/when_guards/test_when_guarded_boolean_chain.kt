// vybe-test: kotlin/when_guards/test_when_guarded_boolean_chain
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun decide(a: Boolean, b: Boolean): String = when {
            a && b -> "both"
            a -> "a"
            b -> "b"
            else -> "none"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((decide(true, true)).toString(), "both")
            __check((decide(true, false)).toString(), "a")
            __check((decide(false, false)).toString(), "none")
        }
