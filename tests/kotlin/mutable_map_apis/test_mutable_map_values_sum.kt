// vybe-test: kotlin/mutable_map_apis/test_mutable_map_values_sum
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            __check((values.values.sum()).toString(), "3")
        }
