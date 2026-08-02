// vybe-test: kotlin/kotlin_list_apis/test_list_map_not_null
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, null, 2, null, 3)
            val out = list.mapNotNull { it?.toString() }.joinToString(",")
            __check((out).toString(), "1,2,3")
        }
