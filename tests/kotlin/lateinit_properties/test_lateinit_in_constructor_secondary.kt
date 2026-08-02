// vybe-test: kotlin/lateinit_properties/test_lateinit_in_constructor_secondary
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Payload {
            lateinit var text: String

            constructor(v: String) {
                text = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Payload("go")
            __check((p.text).toString(), "go")
        }
