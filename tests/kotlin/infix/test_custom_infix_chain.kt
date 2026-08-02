// vybe-test: kotlin/infix/test_custom_infix_chain
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Counter(val base: Int) {
            infix fun plus(other: Int): Int = base + other
            infix fun minus(other: Int): Int = base - other
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter(10)
            __check((c plus 3).toString(), "13")
            __check((c minus 5).toString(), "5")
        }
