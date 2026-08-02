// vybe-test: kotlin/when_guards/test_when_guarded_strings
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = ""
            val out = when {
                text.isEmpty() -> "empty"
                text.length < 2 -> "short"
                else -> "ok"
            }
            __check((out).toString(), "empty")
        }
