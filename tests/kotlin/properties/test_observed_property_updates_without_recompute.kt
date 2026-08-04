// vybe-test: kotlin/properties/test_observed_property_updates_without_recompute
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Item {
            private var count = 0
            val snapshot: Int
                get() = count

            var value: Int
                get() = count
                set(next) { count = next + 1 }
        }

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
            val item = Item()
            item.value = 7
            __p((item.snapshot).toString())
            item.value = 1
            __p((item.snapshot).toString())
            __p((item.value).toString())
        
__check("8\n2\n2")
}
