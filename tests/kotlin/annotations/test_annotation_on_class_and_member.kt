// vybe-test: kotlin/annotations/test_annotation_on_class_and_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated
        class Legacy {
            fun name(): String = "legacy"
        }

        @Suppress("UNUSED_PARAMETER")
        fun tagged(@Deprecated code: Int): String {
            return "tagged"
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
            val legacy = Legacy()
            __p((legacy.name()).toString())
            __p((tagged(1)).toString())
        
__check("legacy\ntagged")
}
