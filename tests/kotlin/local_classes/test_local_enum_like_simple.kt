// vybe-test: kotlin/local_classes/test_local_enum_like_simple
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

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

// Damaged spelling repaired: the enum was declared inside `fun main()`, and
// kotlinc 2.4.10 rejects that — "modifier 'enum' is not applicable to 'local
// class'". Hoisted to top level; measured under kotlinc 2.4.10 it prints the
// expected B.
enum class Mode { A, B, C }

fun main() {
            __p((Mode.B.name).toString())

__check("B")
}
