// vybe-test: kotlin/kotlin_list_apis/test_list_slice_int_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("a", "b", "c", "d")
            val slice = list.slice(IntRange(0, 2))
            __check((slice.joinToString("")).toString(), "abc")
        }
