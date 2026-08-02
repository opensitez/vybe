// vybe-test: kotlin/when_guards/test_when_guarded_by_length
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(s: String): String = when {
            s.length == 0 -> "empty"
            s.length == 1 -> "tiny"
            s.length > 3 -> "long"
            else -> "short"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label("")).toString(), "empty")
            __check((label("a")).toString(), "tiny")
            __check((label("code")).toString(), "long")
        }
