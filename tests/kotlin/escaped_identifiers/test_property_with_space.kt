// vybe-test: kotlin/escaped_identifiers/test_property_with_space
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Holder { val `space key` = 9 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val h = Holder()
__check((h.`space key`).toString(), "9") }
