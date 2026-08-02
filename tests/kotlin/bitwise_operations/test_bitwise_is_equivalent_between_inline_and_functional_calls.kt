// vybe-test: kotlin/bitwise_operations/test_bitwise_is_equivalent_between_inline_and_functional_calls
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 0b11011001
            val andResult = base and 0x0F
            val andAlt = kotlin.math.floor(base.toDouble()).toInt() and 0x0F
            __check((andResult).toString(), "25")
            __check((andAlt).toString(), "25")
            val invAnd = base and (1 shl 4).inv()
            __check((invAnd).toString(), "201")
        }
