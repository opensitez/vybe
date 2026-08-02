// vybe-test: kotlin/kotlin_resource_management/test_resource_in_early_returned_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Slot : AutoCloseable {
            var closed = false
            override fun close() { closed = true }
        }

        fun compute(v: Int): String {
            val slot = Slot()
            slot.use {
                if (v == 0) return "zero"
            }
            return "done"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(0)).toString(), "zero")
        }
