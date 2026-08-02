// vybe-test: kotlin/invoke_operator/test_invoke_tail_style_reuse
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Tail {
            operator fun invoke(v: Int): Tail {
                return if (v <= 0) this else Tail()
            }
            val id: Int = 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tail()
            __check((t(0).id).toString(), "1")
            __check((t(1).id).toString(), "1")
        }
