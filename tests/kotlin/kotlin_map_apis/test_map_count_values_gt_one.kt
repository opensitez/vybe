// vybe-test: kotlin/kotlin_map_apis/test_map_count_values_gt_one
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val overOne = map.filterValues { it > 1 }.size
            val sum = map.values.sum()
            __check((overOne).toString(), "2")
            __check((sum).toString(), "6")
        }
