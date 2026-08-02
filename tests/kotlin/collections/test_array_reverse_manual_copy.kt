// vybe-test: kotlin/collections/test_array_reverse_manual_copy
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            val reversed = Array(nums.size) { idx -> nums[nums.size - 1 - idx] }
            var out = ""
            for (value in reversed) {
                out += value.toString()
            }
            println(out)
            println(reversed[0] + reversed[3])
        }

