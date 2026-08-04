// vybe-test: kotlin/properties/test_property_jvm_field_backing_name_stable_behavior
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Box {
            var value: Int = 0
            val computed: Int
                get() = value + 1
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
            val box = Box()
            box.value = 9
            __p((box.computed).toString())
            box.value = 0
            __p((box.value).toString())
        
__check("10\n0")
}
