// vybe-test: kotlin/mutable_map_apis/test_mutable_map_retain_all_keys
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            values.entries.removeIf { it.key == "b" }
            __check((values.keys.joinToString(",")).toString(), "a,c")
        }
