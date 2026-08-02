// vybe-test: kotlin/collections_iterables/test_list_zip_with_other_list
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf("a", "b", "c")
            val right = listOf(1, 2, 3, 4)
            val pairs = left.zip(right) { l, r -> "$l:$r" }
            __check((pairs.joinToString("|")).toString(), "a:1|b:2|c:3")
        }
