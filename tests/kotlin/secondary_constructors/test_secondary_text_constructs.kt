// vybe-test: kotlin/secondary_constructors/test_secondary_text_constructs
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Text {
            val value: String

            constructor(prefix: String) {
                this.value = prefix
            }

            constructor(prefix: String, suffix: String) : this(prefix) {
                this.value = prefix + suffix
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Text("a").value).toString(), "a")
            __check((Text("a", "b").value).toString(), "ab")
        }
