// vybe-test: kotlin/interfaces/test_interface_property_contract
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Identified {
            val id: Int
        }

        class Item(override val id: Int) : Identified

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Identified = Item(7)
            __check((item.id).toString(), "7")
        }
