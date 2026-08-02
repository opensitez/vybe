// vybe-test: kotlin/secondary_constructors/test_secondary_float_constructor
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Box {
            val value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Double) : this() {
                this.value = v.toInt()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(4.9).value).toString(), "4")
        }
