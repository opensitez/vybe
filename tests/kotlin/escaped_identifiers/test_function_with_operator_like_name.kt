// vybe-test: kotlin/escaped_identifiers/test_function_with_operator_like_name
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun `a b`(x: Int, y: Int): Int = x * y
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`a b`(3, 4)).toString(), "12") }
