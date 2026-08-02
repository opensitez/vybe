// vybe-test: kotlin/kotlin_list_apis/test_list_distinct_by
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("aa", "ab", "bc", "bd", "c")
            val out = list.distinctBy { it.length }
            __check((out.joinToString(",")).toString(), "aa,bc,c")
        }
