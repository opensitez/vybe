// vybe-test: kotlin/kotlin_long_array_apis/test_long_array_size_growth_like_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_long_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = longArrayOf(9L, 8L)
            val b = a.copyOf(3)
            b[2] = 7L
            __check((b.size).toString(), "3")
            __check((b[2]).toString(), "7")
        }
