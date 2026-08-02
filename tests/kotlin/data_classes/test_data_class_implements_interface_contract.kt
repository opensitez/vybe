// vybe-test: kotlin/data_classes/test_data_class_implements_interface_contract
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

interface Identifiable { val id: Int }
        data class Item(override val id: Int, val payload: String) : Identifiable

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Identifiable = Item(7, "payload")
            val a = item.id
            __check((a).toString(), "7")
            __check(((item as Item).payload).toString(), "payload")
        }
