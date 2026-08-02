// vybe-test: kotlin/escaped_identifiers/test_backtick_method_chain
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Counter {
        fun `next value`(x: Int) = x + 1
        fun `next value`(x: Int, y: Int) = x + y
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val c = Counter()
        __check((c.`next value`(3) + c.`next value`(1, 2)).toString(), "7")
    }
