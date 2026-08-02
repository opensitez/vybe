// vybe-test: kotlin/tailrec_functions/test_tailrec_string_reverse_length_based
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun reverseDistance(value: String, idx: Int = 0): Int {
            return if (idx == value.length) 0 else 1 + reverseDistance(value, idx + 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((reverseDistance("kotlin")).toString(), "6")
        }
