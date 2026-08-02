// vybe-test: kotlin/kotlin_closeable_use/test_use_closes_on_exception
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        class Tracker : Closeable {
            var closed = false
            override fun close() { closed = true }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var closed = false
            val tracker = Tracker()
            try {
                tracker.use {
                    __check(("before").toString(), "before")
                    throw IllegalStateException("x")
                }
            } catch (e: Exception) {
                __check((e::class.simpleName).toString(), "IllegalStateException")
                closed = tracker.closed
            }
            __check((closed).toString(), "true")
        }
