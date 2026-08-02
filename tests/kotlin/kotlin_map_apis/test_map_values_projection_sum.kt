// vybe-test: kotlin/kotlin_map_apis/test_map_values_projection_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 10, "b" to 20, "c" to 30)
            val values = map.values.toMutableList()
            values[1] = 25
            __check((values.sum()).toString(), "65")
        }
