// vybe-test: kotlin/mutable_map_apis/test_mutable_map_get_or_default
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            __check((values.getOrDefault("a", 9)).toString(), "1")
            __check((values.getOrDefault("b", 9)).toString(), "9")
        }
