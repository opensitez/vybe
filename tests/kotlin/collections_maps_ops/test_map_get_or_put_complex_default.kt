// vybe-test: kotlin/collections_maps_ops/test_map_get_or_put_complex_default
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("seed" to mutableListOf(1))
            map.getOrPut("seed") { mutableListOf() }.add(2)
            val values = map.getOrPut("fresh") { mutableListOf(9) }
            values.add(10)
            __check((map["seed"]?.size).toString(), "2")
            __check((map["seed"]?.get(1)).toString(), "2")
            __check((map["fresh"]?.size).toString(), "2")
        }
