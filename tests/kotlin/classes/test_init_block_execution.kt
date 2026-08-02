// vybe-test: kotlin/classes/test_init_block_execution
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Counter {
            var count = 0
            init {
                __check(("init").toString(), "init")
            }
            fun increment() {
                count += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            c.increment()
            c.increment()
            __check((c.count).toString(), "2")
        }
