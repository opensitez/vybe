// vybe-test: kotlin/scoping_functions/test_apply_mutates_and_returns_receiver
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = StringBuilder("ko").apply {
                append("t")
                append("l")
                append("in")
            }
            __check((text.toString()).toString(), "kotlin")
            __check((text.length).toString(), "6")
        }
