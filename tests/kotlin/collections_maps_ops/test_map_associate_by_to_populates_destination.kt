// vybe-test: kotlin/collections_maps_ops/test_map_associate_by_to_populates_destination
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("k", "ka", "kb")
            val byLength = mutableMapOf<Int, String>()
            words.associateByTo(byLength, { it.length })
            __check((byLength[1]).toString(), "k")
            __check((byLength[2]).toString(), "kb")
        }
