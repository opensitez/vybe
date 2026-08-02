// vybe-test: kotlin/infix/test_infix_multiple_operator_calls_left_to_right
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class NumberPair(val value: Int) {
            infix fun add(other: NumberPair): NumberPair = NumberPair(this.value + other.value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val total = NumberPair(1) add NumberPair(2) add NumberPair(3)
            __check((total.value).toString(), "6")
        }
