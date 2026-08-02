// vybe-test: kotlin/arrays_ops/test_array_as_list_roundtrip_is_copy
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            val list = nums.toTypedArray().toList()
            val rebuilt = list.toIntArray()
            __check((list.size).toString(), "3")
            __check((rebuilt.joinToString(",")).toString(), "1,2,3")
            __check((list[1]).toString(), "2")
        }
