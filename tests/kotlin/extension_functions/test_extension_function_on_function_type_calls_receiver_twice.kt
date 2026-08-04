// vybe-test: kotlin/extension_functions/test_extension_function_on_function_type_calls_receiver_twice
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun (() -> Int).callTwice(): Int {
            return this() + this()
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
            var state = 0
            val value = { state += 1
state }
            __p((value.callTwice()).toString())
            __p((value.callTwice()).toString())
            __p((state).toString())
        
__check("3\n7\n4")
}
