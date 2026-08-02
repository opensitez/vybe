// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_reference_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = uintArrayOf(1u, 2u)
            val alias = original
            alias[1] = 9u
            __check((original[1].toString()).toString(), "9")
        }
