// vybe-test: kotlin/recursion/test_recursion_sum_map_values
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun sumMap(values: Map<String, Int>): Int {
            if (values.isEmpty()) return 0
            val first = values.entries.first()
            return first.value + sumMap(values - first.key)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumMap(mapOf("a" to 1, "b" to 2))).toString(), "3")
        }
