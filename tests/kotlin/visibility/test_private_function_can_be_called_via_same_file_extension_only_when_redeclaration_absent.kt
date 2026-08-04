// vybe-test: kotlin/visibility/test_private_function_can_be_called_via_same_file_extension_only_when_redeclaration_absent
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            private fun secret(): String = "inside"
        }

        fun itemSecret(item: Item): String = item.access()

        private fun Item.access(): String = secret()

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
            __p((itemSecret(Item())).toString())
        
__check("inside")
}
