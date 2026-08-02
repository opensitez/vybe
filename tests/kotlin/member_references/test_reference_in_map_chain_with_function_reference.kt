// vybe-test: kotlin/member_references/test_reference_in_map_chain_with_function_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Item(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(Item(3), Item(7), Item(9))
            val refs = items.map(Item::value).map { it * 2 }
            __check((refs.joinToString("|")).toString(), "6|14|18")
        }
