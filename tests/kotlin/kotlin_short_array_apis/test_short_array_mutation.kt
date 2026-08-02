// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = shortArrayOf(1, 2)
            data[1] = 9
            __check((data[1].toString()).toString(), "9")
        }
