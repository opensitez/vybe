// vybe-test: kotlin/kotlin_destructuring_maps/test_destructuring_entries_in_filter_step
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val filtered = mapOf("a" to 1, "b" to 2).filter { (k, v) -> k == "b" || v == 1 }
            __check((filtered["a"]).toString(), "1")
            __check((filtered["b"]).toString(), "2")
        }
