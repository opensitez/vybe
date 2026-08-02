// vybe-test: kotlin/ordered_collections/test_map_to_list_round_trip_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("a" to 1, "b" to 2, "c" to 3)
            val map = linkedMapOf<String, Int>()
            map.putAll(list.toMap())
            val rebuilt = map.toList()
            __check((rebuilt.joinToString("|") { "${'$'}{it.first}:${'$'}{it.second}" }).toString(), "a:1|b:2|c:3")
        }
