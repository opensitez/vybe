// vybe-test: kotlin/escaped_identifiers/test_backtick_parameter_name
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun combine(`first part`: Int, `second part`: Int) = `first part` + `second part`
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((combine(4, 5)).toString(), "9") }
