// vybe-test: kotlin/kotlin_resource_management/test_try_catch_still_runs_close
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Token : AutoCloseable {
            var closed = false
            override fun close() {
                closed = true
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val token = Token()
            try {
                throw IllegalStateException("x")
            } catch (_: IllegalStateException) {
                __check(("err").toString(), "err")
            } finally {
                token.close()
            }
            __check((token.closed).toString(), "true")
        }
