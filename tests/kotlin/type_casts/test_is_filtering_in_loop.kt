// vybe-test: kotlin/type_casts/test_is_filtering_in_loop
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

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
            val items: Array<Any?> = arrayOf(1, "x", true, 2.5, null)
            var count = 0
            var stringLen = 0
            var boolSeen = false
            for (item in items) {
                if (item is Int) {
                    count += item
                } else if (item is String) {
                    stringLen = item.length
                } else if (item is Boolean) {
                    boolSeen = item
                }
            }
            __p((count).toString())
            __p((stringLen).toString())
            __p((boolSeen).toString())
        
__check("1\n1\ntrue")
}
