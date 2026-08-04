// vybe-test: kotlin/companion_objects/test_companion_object_can_be_used_as_an_interface_value
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

interface Named {
            fun name(): String
        }

        class Factory {
            companion object : Named {
                override fun name(): String = "factory"
            }
        }

        fun label(source: Named): String = source.name()

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
            val source: Named = Factory.Companion
            __p((label(source)).toString())
            __p((label(Factory.Companion)).toString())
        
__check("factory\nfactory")
}
