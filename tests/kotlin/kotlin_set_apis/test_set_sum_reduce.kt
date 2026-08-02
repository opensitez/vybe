// vybe-test: kotlin/kotlin_set_apis/test_set_sum_reduce
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(2, 4, 6)
            __check((set.sum()).toString(), "12")
            __check((set.reduce { a, b -> a + b }).toString(), "12")
        }
