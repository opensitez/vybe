// vybe-test: kotlin/kotlin_this_qualifiers/test_nested_this_in_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_this_qualifiers.rs

class Box {
            val marker = "box"
            inner class Holder {
                fun show(prefix: String): String {
                    val read = this@Holder
                    return prefix + read.parentName()
                }

                fun parentName(): String {
                    return this@Box.marker
                }
            }
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
            val h = Box().Holder()
            __p((h.show("m=")).toString())
        
__check("m=box")
}
