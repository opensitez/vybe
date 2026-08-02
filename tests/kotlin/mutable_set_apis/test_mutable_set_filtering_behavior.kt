// vybe-test: kotlin/mutable_set_apis/test_mutable_set_filtering_behavior
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            val filtered = values.filter { it > 2 }.toMutableSet()
            __check((filtered.joinToString(",")).toString(), "3,4")
        }
