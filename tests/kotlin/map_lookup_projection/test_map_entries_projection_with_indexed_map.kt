// vybe-test: kotlin/map_lookup_projection/test_map_entries_projection_with_indexed_map
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val indexed = source.entries.mapIndexed { index, e -> "${'$'}{index}:${'$'}{e.key}:${'$'}{e.value}" }
            __check((indexed.joinToString("|")).toString(), "0:a:1|1:b:2|2:c:3")
        }
