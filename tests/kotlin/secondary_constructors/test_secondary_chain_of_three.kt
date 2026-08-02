// vybe-test: kotlin/secondary_constructors/test_secondary_chain_of_three
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Ring {
            val value: Int

            constructor() {
                this.value = 1
            }

            constructor(a: Int) : this() {
                this.value = a
            }

            constructor(a: Int, b: Int) : this(a) {
                this.value = a + b
            }

            constructor(a: Int, b: Int, c: Int) : this(a, b) {
                this.value = this.value + c
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Ring(2, 3).value).toString(), "5")
            __check((Ring(2, 3, 4).value).toString(), "9")
        }
