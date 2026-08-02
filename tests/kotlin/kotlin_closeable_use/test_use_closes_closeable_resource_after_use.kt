// vybe-test: kotlin/kotlin_closeable_use/test_use_closes_closeable_resource_after_use
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        class Tracker : Closeable {
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
            val tracker = Tracker()
            tracker.use {
                __check((it.closed).toString(), "false")
            }
            __check((tracker.closed).toString(), "true")
        }
