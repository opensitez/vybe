// vybe-test: kotlin/list_index_api/test_list_drop_take_boundaries
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            __check((values.take(2).joinToString(",")).toString(), "1,2")
            __check((values.drop(2).joinToString(",")).toString(), "3,4,5")
            __check((values.takeLast(2).joinToString(",")).toString(), "4,5")
            __check((values.dropLast(4).joinToString(",")).toString(), "1")
        }
