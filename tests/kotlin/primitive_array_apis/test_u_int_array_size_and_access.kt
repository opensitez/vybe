// vybe-test: kotlin/primitive_array_apis/test_u_int_array_size_and_access
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = uintArrayOf(1u, 2u, 3u)
            __check((values.size).toString(), "3")
            __check((values[1]).toString(), "2")
        }
