// vybe-test: kotlin/kotlin_class_init_sequences/test_init_block_updates_mutable_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Tracker {
            var count = 0
            init { count = count + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tracker()
            __check((t.count).toString(), "1")
        }
