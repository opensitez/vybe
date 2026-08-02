// vybe-test: kotlin/kotlin_list_filter_apis/test_filter_not_null
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", null, "b", null, "c")
            __check((values.filterNotNull().joinToString(",")).toString(), "a,b,c")
            __check((values.count { it == null }).toString(), "2")
        }
