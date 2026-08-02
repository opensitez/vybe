// vybe-test: kotlin/data_class_copying/test_data_class_copy_in_collections
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Item(val name: String, val count: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(Item("a", 1), Item("b", 2))
            val upgraded = items.map { it.copy(count = it.count + 10) }
            __check((upgraded.joinToString("|") { "${'$'}{it.name}:${'$'}{it.count}" }).toString(), "a:11|b:12")
        }
