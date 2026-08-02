// vybe-test: kotlin/kotlin_map_apis/test_map_to_list_of_pairs
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val pairs = map.toList()
            __check((pairs.size).toString(), "2")
            __check((pairs[0].first).toString(), "a")
            __check((pairs[1].second).toString(), "2")
        }
