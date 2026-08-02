// vybe-test: kotlin/kotlin_sorting_comparators/test_map_indexed_sorting_preserves_indexed_payload
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("c", "a", "b").withIndex().toList().sortedBy { it.value }
            __check((values.map { "${'$'}{it.index}:${'$'}{it.value}" }.joinToString(",")).toString(), "1:a,2:b,0:c")
        }
