// vybe-test: kotlin/escaped_identifiers/test_backtick_top_level_constant
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

val `global value` = 10
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`global value`).toString(), "10") }
