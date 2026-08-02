// vybe-test: kotlin/list_index_api/test_list_element_access_from_reversed_view
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3)
            val rev = values.asReversed()
            __check((rev[0]).toString(), "3")
            __check((rev[1]).toString(), "2")
            __check((rev[2]).toString(), "1")
        }
