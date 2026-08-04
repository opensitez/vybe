// vybe-test: kotlin/collections/test_array_with_homogeneous_structures
// origin: languages/kotlin/tests/kotlin/test_collections.rs

interface Item {
            fun value(): Int
        }

        class NumberItem(val v: Int) : Item {
            override fun value(): Int = v
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
            val boxed: Array<Item> = arrayOf(NumberItem(1), NumberItem(2), NumberItem(3))
            var total = 0
            for (item in boxed) {
                total += item.value()
            }
            __p((total).toString())
            __p((boxed[1].value()).toString())
        
__check("6\n2")
}
