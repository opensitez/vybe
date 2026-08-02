// vybe-test: kotlin/kotlin_list_filter_apis/test_filter_is_instance
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Any> = listOf(1, "a", 2, "b", 3.0)
            val ints = values.filterIsInstance<Int>()
            val strings = values.filterIsInstance<String>()
            __check((ints.joinToString(",")).toString(), "1,2")
            __check((strings.joinToString(",")).toString(), "a,b")
        }
