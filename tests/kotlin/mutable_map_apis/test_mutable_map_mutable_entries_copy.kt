// vybe-test: kotlin/mutable_map_apis/test_mutable_map_mutable_entries_copy
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            val copy = values.toMutableMap()
            copy["a"] = 9
            __check((values["a"]).toString(), "1")
            __check((copy["a"]).toString(), "9")
        }
