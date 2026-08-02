// vybe-test: kotlin/escaped_identifiers/test_backtick_inheritance_override
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

open class Base { open fun `do work`(): Int = 1 }
class Child: Base() { override fun `do work`() = 3 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Child().`do work`()).toString(), "3") }
