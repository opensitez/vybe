// vybe-test: kotlin/data_class_copying/test_data_class_copy_preserves_reference_for_unchanged
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Node(val value: IntArray)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = intArrayOf(1, 2)
            val original = Node(src)
            val copied = original.copy()
            __check((original.value.contentEquals(copied.value)).toString(), "true")
            __check((original.value === copied.value).toString(), "true")
        }
