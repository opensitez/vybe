// vybe-test: kotlin/collections_iterables/test_list_single_and_single_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val only = listOf(42)
            __check((only.single()).toString(), "42")
            __check((listOf<Int>().singleOrNull() ?: -1).toString(), "-1")
        }
