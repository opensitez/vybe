// vybe-test: kotlin/secondary_constructors/test_secondary_noisy_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Log {
            val step: Int

            constructor() {
                this.step = 0
            }

            constructor(v: Int) : this() {
                this.step = v
                __check(("s").toString(), "s")
            }

            constructor(v: Int, extra: Int) : this(v) {
                __check(("e").toString(), "e")
                this.step = v + extra
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Log(2)
            Log(3, 4)
        }
