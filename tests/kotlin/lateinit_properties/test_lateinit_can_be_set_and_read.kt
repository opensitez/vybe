// vybe-test: kotlin/lateinit_properties/test_lateinit_can_be_set_and_read
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Box {
            lateinit var text: String
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box()
            box.text = "k"
            __check((box.text).toString(), "k")
        }
