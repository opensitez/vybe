// vybe-test: kotlin/tailrec_functions/test_tailrec_find_first_non_zero
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun firstNonZero(values: List<Int>, idx: Int = 0): Int {
            return if (idx >= values.size) -1 else if (values[idx] != 0) values[idx] else firstNonZero(values, idx + 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstNonZero(listOf(0, 0, 9, 1))).toString(), "9")
        }
