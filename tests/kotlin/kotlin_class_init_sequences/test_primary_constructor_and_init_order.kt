// vybe-test: kotlin/kotlin_class_init_sequences/test_primary_constructor_and_init_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Counter start {
            val value: Int = start
            init {
                __check((value).toString(), "5")
            }
            constructor(): this(5)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Counter()
        }
