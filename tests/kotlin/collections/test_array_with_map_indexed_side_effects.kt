// vybe-test: kotlin/collections/test_array_with_map_indexed_side_effects
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf("a", "b", "c")
            var marker = ""
            nums.mapIndexed { index, value ->
                marker += index.toString() + value
            }
            __check((marker).toString(), "0a1b2c")
        }
