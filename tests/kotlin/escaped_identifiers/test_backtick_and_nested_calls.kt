// vybe-test: kotlin/escaped_identifiers/test_backtick_and_nested_calls
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class A { fun `outer`(b: String) = b }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((A().`outer`("go")).toString(), "go") }
