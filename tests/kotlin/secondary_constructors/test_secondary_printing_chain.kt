// vybe-test: kotlin/secondary_constructors/test_secondary_printing_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Trace {
            val value: Int

            constructor() {
                this.value = 0
                __check(("zero").toString(), "zero")
            }

            constructor(v: Int) : this() {
                __check((v).toString(), "5")
                this.value = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Trace()
            Trace(5)
        }
