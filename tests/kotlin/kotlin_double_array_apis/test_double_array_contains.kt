// vybe-test: kotlin/kotlin_double_array_apis/test_double_array_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_double_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = doubleArrayOf(1.1, 2.2, 3.3)
            __check((a.contains(2.2)).toString(), "true")
            __check((a.size).toString(), "3")
        }
