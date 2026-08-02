// vybe-test: kotlin/escaped_identifiers/test_backtick_generic_type_alias
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

typealias `String-ID` = String
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val id: `String-ID` = "k"
__check((id).toString(), "k") }
