// vybe-test: kotlin/arrays_ops/test_int_array_with_index_iteration
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val nums = intArrayOf(2, 4, 6)
            var trace = ""
            for ((idx, value) in nums.withIndex()) {
                trace += idx.toString() + ":" + value.toString() + ";"
            }
            println(trace)
        }

