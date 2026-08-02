// vybe-test: kotlin/when_guards/test_when_guarded_null
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = null
            val out = when {
                value == null -> "null"
                value > 0 -> "pos"
                else -> "other"
            }
            __check((out).toString(), "null")
        }
