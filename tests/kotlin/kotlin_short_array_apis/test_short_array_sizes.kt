// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_sizes
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = shortArrayOf(10, 20)
            __check((data.size).toString(), "2")
        }
