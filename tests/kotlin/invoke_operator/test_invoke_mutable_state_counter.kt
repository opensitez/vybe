// vybe-test: kotlin/invoke_operator/test_invoke_mutable_state_counter
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Meter {
            private var total = 0
            operator fun invoke(n: Int) {
                total += n
            }
            fun value(): Int = total
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Meter()
            m(3)
            m(4)
            __check((m.value()).toString(), "7")
        }
