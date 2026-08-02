// vybe-test: kotlin/secondary_constructors/test_secondary_boolean_constructor
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Flag {
            val value: Boolean

            constructor() {
                this.value = false
            }

            constructor(flag: Boolean) : this() {
                this.value = flag
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Flag().value).toString(), "false")
            __check((Flag(true).value).toString(), "true")
        }
