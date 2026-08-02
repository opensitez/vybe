// vybe-test: kotlin/escaped_identifiers/test_backtick_in_when_subject
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val `state value` = "ok"
        val out = when (`state value`) {
            "ok" -> "yes"
            else -> "no"
        }
        __check((out).toString(), "yes")
    }
