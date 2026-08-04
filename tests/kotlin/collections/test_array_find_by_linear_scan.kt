// vybe-test: kotlin/collections/test_array_find_by_linear_scan
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
            val nums = arrayOf(7, 3, 9, 4)
            var found = -1
            var index = 0
            for (value in nums) {
                if (value == 9) {
                    found = value
                    break
                }
                index += 1
            }
            __p((found).toString())
            __p((index).toString())
        
__check("9\n2")
}
