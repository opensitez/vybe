// vybe-test: kotlin/arrays_ops/test_array_for_each_with_indexed_sink
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val nums = intArrayOf(4, 5, 6)
            var out = ""
            nums.forEachIndexed { index, value ->
                out += index.toString() + ":" + value.toString() + ";"
            }
            println(out)
        }

