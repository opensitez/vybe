// vybe-test: kotlin/when_guards/test_when_with_local_boolean
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val enabled = true
            val out = when {
                !enabled -> "off"
                else -> "on"
            }
            __check((out).toString(), "on")
        }
