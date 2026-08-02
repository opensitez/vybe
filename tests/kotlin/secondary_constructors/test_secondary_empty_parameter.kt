// vybe-test: kotlin/secondary_constructors/test_secondary_empty_parameter
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Label {
            val value: String

            constructor(text: String) {
                this.value = text
            }

            constructor() : this("none")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Label().value).toString(), "none")
            __check((Label("yes").value).toString(), "yes")
        }
