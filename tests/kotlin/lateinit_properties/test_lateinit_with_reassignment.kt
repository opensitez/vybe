// vybe-test: kotlin/lateinit_properties/test_lateinit_with_reassignment
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Bag {
            lateinit var label: String

            fun setA() { label = "a" }
            fun setB() { label = "b" }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bag()
            b.setA()
            __check((b.label).toString(), "a")
            b.setB()
            __check((b.label).toString(), "b")
        }
