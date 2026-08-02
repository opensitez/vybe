// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_map_entry_to_list_and_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("x" to 4, "y" to 5)
            val pair = values.entries.map { (k, v) -> Pair(k.length, v) }.toList()
            val sum = pair.sumOf { it.first + it.second }
            __check((sum).toString(), "11")
        }
