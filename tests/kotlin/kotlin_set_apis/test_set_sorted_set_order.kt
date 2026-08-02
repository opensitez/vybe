// vybe-test: kotlin/kotlin_set_apis/test_set_sorted_set_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = sortedSetOf(5, 1, 4, 2, 3)
            __check((set.joinToString(",")).toString(), "1,2,3,4,5")
            __check((set.first()).toString(), "1")
            __check((set.last()).toString(), "5")
        }
