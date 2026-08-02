// vybe-test: kotlin/boolean_logic/test_boolean_array_contains_all
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = booleanArrayOf(true, false, true)
            __check((values.all { it }).toString(), "false")
            __check((values.any { it }).toString(), "true")
            __check((values.count { it }).toString(), "2")
            __check((values.count { !it }).toString(), "1")
        }
