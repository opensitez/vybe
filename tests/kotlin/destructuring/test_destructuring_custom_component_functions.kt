// vybe-test: kotlin/destructuring/test_destructuring_custom_component_functions
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

class Holder(private val a: Int, private val b: Int) {
            operator fun component1() = a
            operator fun component2() = b
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Holder(7, 8)
            val (left, right) = value
            __check((left).toString(), "7")
            __check((right).toString(), "8")
            __check((left + right).toString(), "15")
        }
