// vybe-test: kotlin/interfaces/test_interface_property_requirements
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Identifiable {
            val id: Int
        }

        class Record(override val id: Int) : Identifiable

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r: Identifiable = Record(12)
            __check((r.id).toString(), "12")
        }
