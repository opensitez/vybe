// vybe-test: kotlin/escaped_identifiers/test_interface_with_backtick_member
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

interface `I-Thing` { val `prop value`: Int }
class C: `I-Thing` { override val `prop value` = 11 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().`prop value`).toString(), "11") }
