// vybe-test: kotlin/escaped_identifiers/test_nested_backtick_class
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Outer {
        inner class `Inner Type`(val `count value`: Int)
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        __check((Outer().`Inner Type`(7).`count value`).toString(), "7")
    }
