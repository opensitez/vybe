// vybe-test: kotlin/kotlin_double_array_apis/test_double_array_average
// origin: languages/kotlin/tests/kotlin/test_kotlin_double_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = doubleArrayOf(2.0, 4.0, 6.0)
            __check((a.average()).toString(), "4.0")
        }
