// vybe-test: kotlin/kotlin_map_apis/test_map_filter_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("aa" to 1, "bb" to 2, "abc" to 3)
            val filtered = map.filterKeys { it.length > 2 }
            __check((filtered.size).toString(), "1")
            __check((filtered["abc"]).toString(), "3")
        }
