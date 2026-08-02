// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_copy
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = shortArrayOf(1, 2)
            val b = a.copyOf()
            b[0] = 10
            __check((a[0].toString()).toString(), "1")
            __check((b[0].toString()).toString(), "10")
        }
