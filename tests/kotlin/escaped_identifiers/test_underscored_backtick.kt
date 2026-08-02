// vybe-test: kotlin/escaped_identifiers/test_underscored_backtick
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val `a b_c` = 4
__check((`a b_c`).toString(), "4") }
