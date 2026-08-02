// vybe-test: kotlin/collections_maps_ops/test_map_replace_overwrites_only_when_present
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            __check((map.replace("a", 7)).toString(), "1")
            __check((map.replace("c", 9)).toString(), "null")
            __check((map["a"]).toString(), "7")
            __check((map["c"] ?: -1).toString(), "-1")
        }
