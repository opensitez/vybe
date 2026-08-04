// vybe-test: kotlin/object_declarations/test_object_implements_multiple_interfaces
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Named {
            fun name(): String
        }

        interface Versioned {
            fun version(): Int
        }

        object Metadata : Named, Versioned {
            override fun name(): String = "meta"
            override fun version(): Int = 1
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
            __p((Metadata.name()).toString())
            __p((Metadata.version()).toString())
        
__check("meta\n1")
}
