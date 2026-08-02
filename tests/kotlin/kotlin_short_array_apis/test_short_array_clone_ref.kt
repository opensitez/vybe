// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_clone_ref
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = shortArrayOf(2, 4, 6)
            val b = a
            b[1] = 9
            __check((a[1].toString()).toString(), "9")
        }
