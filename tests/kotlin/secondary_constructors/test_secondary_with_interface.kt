// vybe-test: kotlin/secondary_constructors/test_secondary_with_interface
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

interface Marker

        class D : Marker {
            val value: Int

            constructor() {
                this.value = 1
            }

            constructor(v: Int) : this() {
                this.value = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((D().value).toString(), "1")
            __check((D(8).value).toString(), "8")
        }
