// vybe-test: kotlin/kotlin_long_array_apis/test_long_array_literals_and_arith
// origin: languages/kotlin/tests/kotlin/test_kotlin_long_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = longArrayOf(1L, 2L, 3L)
            __check((a[2] - a[0]).toString(), "2")
        }
