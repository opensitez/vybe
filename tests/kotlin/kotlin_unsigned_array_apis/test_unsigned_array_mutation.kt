// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = uintArrayOf(1u, 2u)
            val b = ubyteArrayOf(1u, 2u)
            u[0] = 9u
            b[1] = 8u
            __check((u[0].toString()).toString(), "9")
            __check((b[1].toString()).toString(), "8")
        }
