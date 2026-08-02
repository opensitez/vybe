// vybe-test: kotlin/scoping_functions/test_apply_with_nested_with_reads_its_properties
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Item {
            val id = 1
            val prefix = "item"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = Item().apply {
                id.toString()
            }.let { it.id }
            __check((text).toString(), "1")
            val withText = with(Item()) { "$prefix-$id" }
            __check((withText).toString(), "item-1")
        }
