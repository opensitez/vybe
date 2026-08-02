// vybe-test: kotlin/kotlin_destructuring_maps/test_map_entry_destructure_with_default_for_missing
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("a" to 1)
            val (a, b) = values.entries.first()
            __check((a).toString(), "a")
            __check((values[a] ?: b).toString(), "1")
        }
