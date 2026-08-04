// vybe-test: kotlin/collections/test_array_range_for_indices_bounds
// origin: languages/kotlin/tests/kotlin/test_collections.rs

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5)
            var evenTotal = 0
            for (i in nums.indices) {
                if (i % 2 == 1) evenTotal += nums[i]
            }
            __p((evenTotal).toString())
            __p((nums.lastIndex).toString())
        
__check("6\n4")
}
