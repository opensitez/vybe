// vybe-test: kotlin/infix/test_infix_with_if_expression
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class IntPair(val first: Int, val second: Int) {
            infix fun merge(other: IntPair): Int {
                return (first + second) + (other.first + other.second)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = IntPair(1, 2)
            val b = IntPair(3, 4)
            __check((a merge b).toString(), "10")
            __check((a.first + a.second).toString(), "3")
            __check((b.first + b.second).toString(), "7")
        }
