// vybe-test: kotlin/type_cast_edges/test_is_works_with_subtype
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

interface Animal { fun kind(): String }
        class Cat : Animal { override fun kind() = "cat" }
        class Dog : Animal { override fun kind() = "dog" }

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
            val a: Animal = Cat()
            __p((a is Animal).toString())
            __p((a is Cat).toString())
            __p((a is Dog).toString())
            __p((a.kind()).toString())
        
__check("true\ntrue\nfalse\ncat")
}
