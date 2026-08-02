// vybe-test: kotlin/secondary_constructors/test_secondary_counter_reset
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Seq {
            val n: Int

            constructor() {
                this.n = 0
            }

            constructor(v: Int) : this() {
                this.n = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Seq().n).toString(), "0")
            __check((Seq(12).n).toString(), "12")
        }
