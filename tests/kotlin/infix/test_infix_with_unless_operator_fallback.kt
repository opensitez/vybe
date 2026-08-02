// vybe-test: kotlin/infix/test_infix_with_unless_operator_fallback
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Counter(val value: Int) {
            infix fun plus(other: Counter): Int = this.value + other.value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Counter(4)
            val right = Counter(6)
            __check((left plus right).toString(), "10")
            __check((right plus left).toString(), "10")
        }
