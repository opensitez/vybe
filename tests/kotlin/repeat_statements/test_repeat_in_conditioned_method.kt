// vybe-test: kotlin/repeat_statements/test_repeat_in_conditioned_method
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

class Repeater {
            fun build(a: Int): Int {
                var sum = 0
                repeat(a) { i -> sum += i }
                return sum
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Repeater().build(6)).toString(), "15")
        }
