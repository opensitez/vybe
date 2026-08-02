// vybe-test: kotlin/iterator_protocol/test_iterator_over_map_entries
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val values = map.entries.joinToString("|") { it.key + ":" + it.value }
            __check((values).toString(), "a:1|b:2")
            __check((map.entries.size).toString(), "2")
        }
