// vybe-test: kotlin/primitive_array_apis/test_ubyte_array_to_ints
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = ubyteArrayOf(250u, 1u)
            val first = values[0].toInt()
            __check((first).toString(), "250")
            __check((values.joinToString(",")).toString(), "250,1")
        }
