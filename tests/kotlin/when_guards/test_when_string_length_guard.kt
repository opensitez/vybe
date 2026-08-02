// vybe-test: kotlin/when_guards/test_when_string_length_guard
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "kotlin"
            val out = when {
                s.isEmpty() -> "empty"
                s.length < 3 -> "tiny"
                s.length == 6 -> "size6"
                else -> "other"
            }
            __check((out).toString(), "size6")
        }
