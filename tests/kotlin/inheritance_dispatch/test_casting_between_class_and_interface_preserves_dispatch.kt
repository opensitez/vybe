// vybe-test: kotlin/inheritance_dispatch/test_casting_between_class_and_interface_preserves_dispatch
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Labeled {
            fun label(): String = "interface"
        }

        open class Base {
            open fun label(): String = "base"
        }

        class Widget : Base(), Labeled {
            override fun label(): String = "widget"
        }

        fun callThroughInterface(value: Labeled): String {
            return value.label()
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
            val root: Base = Widget()
            val viaInterface = root as Labeled
            __p(((root as Widget).label()).toString())
            __p((callThroughInterface(viaInterface)).toString())
        
__check("widget\nwidget")
}
