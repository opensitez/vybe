// vybe-test: kotlin/range_apis/test_range_of_indices_used_with_list
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = listOf(10, 20, 30, 40)
            __check((data.indices.toList().joinToString(",")).toString(), "0,1,2,3")
            __check((data.indices.last).toString(), "3")
            __check((data[ data.indices.last ]).toString(), "40")
        }
