// vybe-test: kotlin/secondary_constructors/test_secondary_with_super_base
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

open class P(val id: Int)

        class C : P {
            val tag: Int

            constructor() : super(1) {
                this.tag = 2
            }

            constructor(multiplier: Int) : super(multiplier) {
                this.tag = multiplier * 2
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C().id).toString(), "1")
            __check((C(3).id).toString(), "3")
            __check((C(3).tag).toString(), "6")
        }
