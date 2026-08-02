// vybe-test: kotlin/kotlin_map_apis/test_map_sum_values_of_mutable_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 2, "b" to 4)
            __check(((map["a"] ?: 0) + (map["b"] ?: 0)).toString(), "6")
            map["a"] = map["a"]!! + 3
            __check((map["a"]).toString(), "5")
        }
