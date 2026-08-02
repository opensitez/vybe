// vybe-test: kotlin/throwing_recovery/test_throw_and_finalize
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

class Holder {
            fun close() = __check(("closed").toString(), "closed")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            try {
                throw Exception("x")
            } finally {
                h.close()
            }
        }
