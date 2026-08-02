// vybe-test: kotlin/kotlin_list_apis/test_list_flat_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(listOf(1, 2), listOf(3, 4))
            val out = list.flatMap { it.map { n -> n * 2 } }
            __check((out.joinToString(",")).toString(), "2,4,6,8")
        }
