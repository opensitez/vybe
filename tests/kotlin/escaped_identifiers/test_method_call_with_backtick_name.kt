// vybe-test: kotlin/escaped_identifiers/test_method_call_with_backtick_name
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Api { fun `compute next`(x: Int) = x + 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = Api()
__check((a.`compute next`(5)).toString(), "7") }
