// vybe-test: kotlin/escaped_identifiers/test_reserved_name_in_lambda
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val `if` = { x: Int -> x + 1 }
        __check((`if`(2)).toString(), "3")
    }
