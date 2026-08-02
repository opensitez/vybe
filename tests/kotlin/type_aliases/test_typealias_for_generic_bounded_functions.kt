// vybe-test: kotlin/type_aliases/test_typealias_for_generic_bounded_functions
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias ComparableList<T> = List<T>

        fun <T : Comparable<T>> maxOfList(values: ComparableList<T>): T {
            return values.maxOrNull()!!
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words: ComparableList<String> = listOf("bb", "aaa", "c")
            __check((maxOfList(words)).toString(), "c")
        }
