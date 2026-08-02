// vybe-test: kotlin/kotlin_closeable_use/test_use_returns_lambda_result
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        class Tracker : Closeable {
            var closed = false
            override fun close() {
                closed = true
            }
            fun value() = "done"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tracker = Tracker()
            val out = tracker.use { t ->
                t.value()
            }
            __check((out).toString(), "done")
            __check((tracker.closed).toString(), "true")
        }
