// vybe-test: kotlin/kotlin_list_apis/test_list_slice_simple
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            val slice = list.slice(1..3)
            __check((slice.joinToString(",")).toString(), "2,3,4")
        }
