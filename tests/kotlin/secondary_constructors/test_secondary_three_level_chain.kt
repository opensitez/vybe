// vybe-test: kotlin/secondary_constructors/test_secondary_three_level_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class C {
            val value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }

            constructor(v: Int, extra: Int) : this(v) {
                this.value = v + extra
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C().value).toString(), "0")
            __check((C(4).value).toString(), "4")
            __check((C(4, 5).value).toString(), "9")
        }
