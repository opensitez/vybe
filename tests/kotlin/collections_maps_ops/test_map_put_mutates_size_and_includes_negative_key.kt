// vybe-test: kotlin/collections_maps_ops/test_map_put_mutates_size_and_includes_negative_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf<Int, Int>()
            map.put(-1, 10)
            map[-2] = 20
            __check((map.size).toString(), "2")
            __check((map[-1] + map[-2]).toString(), "30")
        }
