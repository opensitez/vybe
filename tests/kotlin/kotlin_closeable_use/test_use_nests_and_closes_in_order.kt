// vybe-test: kotlin/kotlin_closeable_use/test_use_nests_and_closes_in_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        val events = StringBuilder()

        class Tracker(val tag: String) : Closeable {
            override fun close() {
                events.append(tag)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Tracker("a").use {
                it
                Tracker("b").use {
                    it
                }
            }
            __check((events.toString()).toString(), "ba")
        }
