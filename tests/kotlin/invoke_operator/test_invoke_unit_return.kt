// vybe-test: kotlin/invoke_operator/test_invoke_unit_return
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Sink {
            private var logged = false
            operator fun invoke(v: Int): Unit { logged = v > 0 }
            fun status() = logged
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Sink()
            s(3)
            __check((s.status()).toString(), "true")
            s(-1)
            __check((s.status()).toString(), "false")
        }
