// vybe-test: kotlin/boolean_logic/test_boolean_join_to_string_and_join
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flags = booleanArrayOf(true, false, true)
            val text = flags.joinToString("|") { it.toString() }
            __check((text).toString(), "true|false|true")
            __check((text.length).toString(), "14")
        }
