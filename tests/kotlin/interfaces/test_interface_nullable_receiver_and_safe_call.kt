// vybe-test: kotlin/interfaces/test_interface_nullable_receiver_and_safe_call
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Reporter {
            fun report(): String
        }

        class Logger : Reporter {
            override fun report(): String = "ok"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val good: Reporter? = Logger()
            val bad: Reporter? = null
            __check((good?.report() ?: "missing").toString(), "ok")
            __check((bad?.report() ?: "missing").toString(), "missing")
        }
