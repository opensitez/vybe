// vybe-test: kotlin/kotlin_set_apis/test_set_minus_operator
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = linkedSetOf(1, 2, 3, 4)
            val b = a - listOf(2, 4)
            __check((b.size).toString(), "2")
            __check((b.joinToString(",")).toString(), "1,3")
        }
