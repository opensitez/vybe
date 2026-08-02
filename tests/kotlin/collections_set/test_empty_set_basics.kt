// vybe-test: kotlin/collections_set/test_empty_set_basics
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = emptySet<Int>()
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
            __check((values.contains(1)).toString(), "false")
        }
