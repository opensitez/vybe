// vybe-test: kotlin/classes/test_class_secondary_constructor_delegates_to_primary
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Box {
            val value: Int
            val label: String

            constructor(value: Int) : this(value, "default") {
                __p(("secondary").toString())
            }

            constructor(value: Int, label: String) {
                this.value = value
                this.label = label
            }

            fun describe(): String {
                return label + ":" + value
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
            val item = Box(5)
            __p((item.describe()).toString())
        
__check("secondary\ndefault:5")
}
