// vybe-test: kotlin/arrays_ops/test_array_of_nulls_defaults_to_null_and_can_mutate_indices
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val slots = arrayOfNulls<String>(3)
            val before = slots.joinToString(",") { it ?: "null" }
            slots[1] = "value"
            val after = slots.joinToString(",") { it ?: "null" }
            __check((before).toString(), "null,null,null")
            __check((after).toString(), "null,value,null")
            __check((slots.count { it == null }).toString(), "2")
        }
