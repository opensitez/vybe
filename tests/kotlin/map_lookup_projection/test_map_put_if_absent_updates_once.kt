// vybe-test: kotlin/map_lookup_projection/test_map_put_if_absent_updates_once
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = linkedMapOf("a" to 1)
            val existing = source.putIfAbsent("a", 99)
            val added = source.putIfAbsent("b", 2)
            __check((existing).toString(), "1")
            __check((added).toString(), "null")
            __check((source["a"]).toString(), "1")
            __check((source["b"]).toString(), "2")
        }
