// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_copy_of
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = uintArrayOf(2u, 4u, 6u)
            val copy = u.copyOf()
            copy[0] = 99u
            __check((u[0].toString()).toString(), "2")
            __check((copy[0].toString()).toString(), "99")
        }
