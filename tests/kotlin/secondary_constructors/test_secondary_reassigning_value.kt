// vybe-test: kotlin/secondary_constructors/test_secondary_reassigning_value
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Counter {
            var value: Int

            constructor() {
                this.value = 1
            }

            constructor(value: Int, double: Boolean) : this(value) {
                if (double) {
                    this.value = value * 2
                } else {
                    this.value = value
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter(3, true).value).toString(), "6")
            __check((Counter(3, false).value).toString(), "3")
        }
