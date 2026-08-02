// vybe-test: kotlin/data_classes/test_data_class_array_is_equal_by_structural_values
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Record(val values: IntArray)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Record(intArrayOf(1, 2))
            val b = Record(intArrayOf(1, 2))
            __check((a == b).toString(), "false")
            __check((a.values.contentToString()).toString(), "[1, 2]")
        }
