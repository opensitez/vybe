// vybe-test: kotlin/block_expressions/test_block_in_class_initializer
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

class K {
        val x = run {
            val a = 1
            val b = 2
            a + b
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().x).toString(), "3") }
