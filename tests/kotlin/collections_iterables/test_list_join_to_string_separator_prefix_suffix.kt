// vybe-test: kotlin/collections_iterables/test_list_join_to_string_separator_prefix_suffix
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names = listOf("a", "b", "c")
            __check((names.joinToString(prefix = "[", postfix = "]", separator = ",")).toString(), "[a,b,c]")
        }
