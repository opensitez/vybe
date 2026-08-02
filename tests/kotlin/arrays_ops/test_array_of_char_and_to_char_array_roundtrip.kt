// vybe-test: kotlin/arrays_ops/test_array_of_char_and_to_char_array_roundtrip
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kot"
            val chars = text.toCharArray()
            __check((chars.joinToString(",")).toString(), "k,o,t")
            __check((String(chars)).toString(), "kot")
        }
