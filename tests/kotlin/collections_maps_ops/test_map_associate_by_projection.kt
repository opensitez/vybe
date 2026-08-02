// vybe-test: kotlin/collections_maps_ops/test_map_associate_by_projection
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("one", "two", "three")
            val map = words.associateBy { it.first() }
            __check((map['o']).toString(), "one")
            __check((map['t']).toString(), "three")
        }
