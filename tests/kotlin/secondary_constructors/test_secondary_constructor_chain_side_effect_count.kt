// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_chain_side_effect_count
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class SequenceTracker {
            val value: Int

            constructor() {
                __check(("base").toString(), "base")
                this.value = 0
            }

            constructor(start: Int) : this() {
                __check(("fromStart").toString(), "0")
                this.value = start
            }

            constructor(start: Int, step: Int) : this(start) {
                __check(("fromStep").toString(), "fromStart")
                this.value = start + step
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((SequenceTracker().value).toString(), "3")
            __check((SequenceTracker(3).value).toString(), "fromStep")
            __check((SequenceTracker(3, 4).value).toString(), "7")
        }
