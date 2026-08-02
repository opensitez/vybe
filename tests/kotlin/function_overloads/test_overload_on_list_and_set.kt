// vybe-test: kotlin/function_overloads/test_overload_on_list_and_set
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun size(items: List<Int>): Int = items.size
        fun size(items: Set<Int>): Int = items.size + 10
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((size(listOf(1, 2, 3))).toString(), "3")
            __check((size(setOf(1, 2, 3))).toString(), "13")
        }
