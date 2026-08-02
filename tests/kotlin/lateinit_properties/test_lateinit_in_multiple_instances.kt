// vybe-test: kotlin/lateinit_properties/test_lateinit_in_multiple_instances
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Holder {
            lateinit var note: String
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Holder()
            val b = Holder()
            a.note = "A"
            b.note = "B"
            __check((a.note).toString(), "A")
            __check((b.note).toString(), "B")
        }
