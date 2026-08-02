// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_slice_windowed
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(10, 3, 5, 7).sorted()
            val windows = values.windowed(2)
            __check((windows.size).toString(), "3")
            __check((windows[0].joinToString(",")).toString(), "1,3")
            __check((windows[1].joinToString(",")).toString(), "3,5")
        }
