// vybe-test: kotlin/kotlin_destructuring_maps/test_entry_to_map_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "bb" to 2)
            val out = map
                .toList()
                .associate { (k, v) -> Pair(k + v, v + 1) }
            __check((out["a1"]).toString(), "2")
            __check((out["bb2"]).toString(), "3")
        }
