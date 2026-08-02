// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_reassigns_var_property
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Box {
            var value: Int

            constructor() {
                this.value = 10
            }

            constructor(input: Int) : this() {
                this.value += input
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().value).toString(), "10")
            __check((Box(5).value).toString(), "15")
        }
