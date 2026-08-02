// vybe-test: kotlin/kotlin_set_apis/test_set_empty_union
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf<Int>().union(setOf(1, 2, 2))
            __check((set.size).toString(), "2")
            __check((set.joinToString(",")).toString(), "1,2")
        }
