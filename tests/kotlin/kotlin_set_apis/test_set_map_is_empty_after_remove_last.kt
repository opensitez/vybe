// vybe-test: kotlin/kotlin_set_apis/test_set_map_is_empty_after_remove_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1)
            set.remove(1)
            __check((set.isEmpty()).toString(), "true")
        }
