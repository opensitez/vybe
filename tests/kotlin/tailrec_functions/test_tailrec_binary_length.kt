// vybe-test: kotlin/tailrec_functions/test_tailrec_binary_length
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun binaryLen(values: IntArray, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= values.size) acc else binaryLen(values, idx + 1, acc + values[idx])
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((binaryLen(intArrayOf(1, 2, 3, 4))).toString(), "10")
        }
