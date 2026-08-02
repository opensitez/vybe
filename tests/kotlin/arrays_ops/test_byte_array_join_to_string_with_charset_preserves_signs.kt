// vybe-test: kotlin/arrays_ops/test_byte_array_join_to_string_with_charset_preserves_signs
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, -2, 127, -128)
            __check((bytes.joinToString("|")).toString(), "1,-2,127,-128")
            val first = bytes[1].toInt()
            val last = bytes[3].toInt()
            __check((first + last).toString(), "-130")
        }
