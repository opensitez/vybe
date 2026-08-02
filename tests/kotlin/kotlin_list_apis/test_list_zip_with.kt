// vybe-test: kotlin/kotlin_list_apis/test_list_zip_with
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf("a", "b", "c")
            val b = listOf(1, 2, 3, 4)
            __check((a.zip(b).joinToString(",") { it.first + it.second.toString() }).toString(), "a1,b2,c3")
        }
