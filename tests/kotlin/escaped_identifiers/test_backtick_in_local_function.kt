// vybe-test: kotlin/escaped_identifiers/test_backtick_in_local_function
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        fun `local op`(x: Int, y: Int) = x * y
        __check((`local op`(2, 6)).toString(), "12")
    }
