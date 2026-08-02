// vybe-test: kotlin/kotlin_map_apis/test_map_filter_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 4, "c" to 2)
            val filtered = map.filterValues { it >= 3 }
            __check((filtered.size).toString(), "1")
            __check((filtered["b"]).toString(), "4")
            __check((filtered.containsKey("a")).toString(), "false")
        }
