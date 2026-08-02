// vybe-test: kotlin/kotlin_list_apis/test_list_fold_seed
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3)
            __check((list.fold("start") { acc, value -> acc + value.toString() }).toString(), "start123")
        }
