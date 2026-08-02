// vybe-test: kotlin/secondary_constructors/test_secondary_computed_value
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Compute {
            val value: Int

            constructor(base: Int) {
                this.value = base * 2
            }

            constructor(base: Int, factor: Int) : this(base) {
                this.value = base + factor
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Compute(4, 5)
            __check((c.value).toString(), "9")
        }
