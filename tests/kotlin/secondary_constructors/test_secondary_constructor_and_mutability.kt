// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_and_mutability
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Counter {
            var value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }

            constructor(v: Int, inc: Int, dec: Int) : this(v, inc) {
                this.value = this.value + inc - dec
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter(5, 1, 1).value).toString(), "5")
        }
