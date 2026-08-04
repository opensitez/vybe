// vybe-test: kotlin/sealed_types/test_sealed_class_with_object_leaf
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Option {
            object Empty : Option()
            class Value(val value: Int) : Option()
        }

        fun label(value: Option): String {
            return when (value) {
                is Option.Empty -> "empty"
                is Option.Value -> value.value.toString()
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
            __p((label(Option.Empty)).toString())
            __p((label(Option.Value(5))).toString())
        
__check("empty\n5")
}
