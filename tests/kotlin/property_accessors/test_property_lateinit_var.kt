// vybe-test: kotlin/property_accessors/test_property_lateinit_var
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Holder {
            lateinit var text: String
            fun run() {
                text = "ok"
                __check((text).toString(), "ok")
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Holder().run()
        }
