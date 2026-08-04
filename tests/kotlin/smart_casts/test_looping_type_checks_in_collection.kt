// vybe-test: kotlin/smart_casts/test_looping_type_checks_in_collection
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

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
            val values = listOf<Any>(1, "two", 3, "four")
            var strings = 0
            var totalLen = 0
            for (item in values) {
                if (item is String) {
                    strings += 1
                    totalLen += item.length
                }
            }
            __p((strings).toString())
            __p((totalLen).toString())
        
__check("2\n7")
}
