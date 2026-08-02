// vybe-test: kotlin/escaped_identifiers/test_backtick_constructor_parameter
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Box(val `label text`: String)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b = Box("x")
__check((b.`label text`).toString(), "x") }
