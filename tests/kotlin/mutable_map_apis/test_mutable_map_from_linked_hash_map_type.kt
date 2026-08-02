// vybe-test: kotlin/mutable_map_apis/test_mutable_map_from_linked_hash_map_type
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.LinkedHashMap<String, Int>()
            values["a"] = 1
            values["b"] = 2
            __check((values["a"]).toString(), "1")
            __check((values.keys.joinToString(",")).toString(), "a,b")
        }
