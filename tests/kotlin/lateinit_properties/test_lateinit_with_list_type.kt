// vybe-test: kotlin/lateinit_properties/test_lateinit_with_list_type
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Collector {
            lateinit var values: MutableList<Int>
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Collector()
            c.values = mutableListOf(1, 2)
            c.values.add(3)
            __check((c.values.joinToString(",")).toString(), "1,2,3")
        }
