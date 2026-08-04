// vybe-test: kotlin/inheritance_dispatch/test_interface_dispatch_is_polymorphic
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Reader {
            fun read(): String
        }

        class A : Reader {
            override fun read(): String = "a"
        }

        class B : Reader {
            override fun read(): String = "b"
        }

        fun emit(readers: Array<Reader>): String {
            var total = ""
            for (reader in readers) {
                total += reader.read()
            }
            return total
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
            __p((emit(arrayOf(A(), B()))).toString())
        
__check("ab")
}
