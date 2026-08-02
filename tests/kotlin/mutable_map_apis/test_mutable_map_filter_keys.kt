// vybe-test: kotlin/mutable_map_apis/test_mutable_map_filter_keys
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val filtered = values.filterValues { it > 1 }
            __check((filtered.joinToString(",") { it.key + ":" + it.value }).toString(), "b:2,c:3")
        }
