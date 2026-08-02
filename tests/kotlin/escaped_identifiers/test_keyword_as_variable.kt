// vybe-test: kotlin/escaped_identifiers/test_keyword_as_variable
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val `class` = 7
__check((`class`).toString(), "7") }
