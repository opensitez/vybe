// vybe-test: kotlin/visibility/test_private_setter_does_not_block_reading_property
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Counter {
            var value: Int = 4
                private set

            fun increment() {
                value += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.increment()
            __check((counter.value).toString(), "5")
            counter.increment()
            __check((counter.value).toString(), "6")
        }
