// vybe-test: kotlin/kotlin_set_apis/test_set_max_min
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = sortedSetOf(10, 5, 12, 7)
            __check((set.minOrNull()).toString(), "5")
            __check((set.maxOrNull()).toString(), "12")
        }
