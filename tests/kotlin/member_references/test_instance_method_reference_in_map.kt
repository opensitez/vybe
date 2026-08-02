// vybe-test: kotlin/member_references/test_instance_method_reference_in_map
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Item(val value: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(Item("a"), Item("b"))
            val labels = items.map(Item::value).joinToString("|")
            __check((labels).toString(), "a|b")
        }
