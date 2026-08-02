// vybe-test: kotlin/raw_strings/test_raw_string_with_dollar_but_no_expr
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 1
            val text = """$${x}"""
            __check((text).toString(), "\$1")
        }
