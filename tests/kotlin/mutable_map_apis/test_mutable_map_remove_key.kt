// vybe-test: kotlin/mutable_map_apis/test_mutable_map_remove_key
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            val removed = values.remove("a")
            __check((removed).toString(), "1")
            __check((values.size).toString(), "1")
        }
